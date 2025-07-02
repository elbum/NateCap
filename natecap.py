from pywinauto import Application
import time
import sys
import psutil
import hashlib
import sqlite3
from datetime import datetime

TARGET_PROCESS = 'NateOnBiz.exe'

def extract_message_text(msg_html_view):
    def collect_texts(ctrl):
        texts = []
        try:
            children = ctrl.children()
            if not children:
                # 마지막(leaf)이면서 Text 타입만 수집
                if ctrl.element_info.control_type == "Text":
                    t = ctrl.window_text()
                    if t:
                        texts.append(t)
            else:
                for c in children:
                    texts.extend(collect_texts(c))
        except Exception:
            pass
        return texts
    all_lines = collect_texts(msg_html_view)
    return "\n".join(all_lines)

def get_message_key(message):
    """쪽지 본문으로부터 SHA-256 해시값을 생성하여 유니크 키로 사용"""
    return hashlib.sha256(message.encode("utf-8")).hexdigest()

def init_db(db_path="natecap_messages.db"):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS messages (
            key TEXT PRIMARY KEY,
            created_at TEXT NOT NULL,
            user TEXT NOT NULL,
            content TEXT NOT NULL
        )
    ''')
    conn.commit()
    return conn

def main():
    print(f"'{TARGET_PROCESS}' 쪽지창 메시지 추출(전체 탐색)을 시작합니다...")
    printed_keys = set()
    conn = init_db()
    c = conn.cursor()
    while True:
        try:
            target_pids = [p.pid for p in psutil.process_iter(['name']) if p.info['name'] and p.info['name'].lower() == TARGET_PROCESS.lower()]
            for pid in target_pids:
                try:
                    app = Application(backend="uia").connect(process=pid)
                    windows = app.windows()
                    for win in windows:
                        # 모든 윈도우의 타이틀/클래스명 출력
                        # print(f"윈도우 타이틀: {win.window_text()}, 클래스명: {win.element_info.class_name}")
                        # 클래스명이 #32770 인 경우에만 쪽지 본문 추출
                        if win.element_info.class_name == "#32770":
                            print(f"클래스명이 #32770 인 윈도우: {win.window_text()}")

                            try:
                                message = extract_message_text(win)
                                if message:
                                    key = get_message_key(message.strip())
                                    # DB에 이미 저장된 키인지 확인
                                    c.execute("SELECT 1 FROM messages WHERE key=?", (key,))
                                    exists = c.fetchone()
                                    if not exists:
                                        created_at = datetime.now().isoformat(sep=' ', timespec='seconds')
                                        user = win.window_text()
                                        # '님의 쪽지'가 포함되어 있으면 앞부분만 추출
                                        if user.endswith("님의 쪽지"):
                                            user = user[:user.rfind("님의 쪽지")].strip()
                                        # user 값이 없거나 '쪽지 쓰기'인 경우는 저장하지 않음
                                        if not user or user == "쪽지 쓰기":
                                            continue
                                        c.execute(
                                            "INSERT INTO messages (key, created_at, user, content) VALUES (?, ?, ?, ?)",
                                            (key, created_at, user, message)
                                        )
                                        conn.commit()
                                        print("\n===== 쪽지 본문 (DB 저장됨) =====")
                                        print(f"[KEY] {key}")
                                        print(f"[생성시각] {created_at}")
                                        print(f"[USER] {user}")
                                        print(message)
                                        print("============================\n")
                                        printed_keys.add(key)
                            except Exception as e:
                                print(f"MsgHtmlView 탐색 실패: {e}")
                        else:
                            continue
                    
                except Exception as e:
                    print(f"PID {pid} 연결 실패: {e}")
            time.sleep(5)
        except KeyboardInterrupt:
            print("\n프로그램을 종료합니다.")
            conn.close()
            sys.exit(0)

if __name__ == "__main__":
    main()