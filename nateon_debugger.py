from pywinauto import Desktop, Application
import time
import sys
import psutil
import hashlib

TARGET_PROCESS = 'NateOnBiz.exe'


def print_window_info(window, indent=0, printed_ids=None):
    if printed_ids is None:
        printed_ids = set()
    prefix = '  ' * indent
    try:
        # 중복 출력 방지: 이미 출력한 컨트롤은 건너뜀
        rid = getattr(window.element_info, 'runtime_id', None)
        if rid is not None:
            rid_tuple = tuple(rid) if isinstance(rid, list) else rid
            if rid_tuple in printed_ids:
                return
            printed_ids.add(rid_tuple)
        # 모든 컨트롤의 텍스트를 한 줄씩 출력
        text = window.window_text()
        if text:
            print(f"{prefix}[텍스트] {text}")
        print(f"{prefix}윈도우 타이틀: {window.window_text()}")
        print(f"{prefix}클래스명: {window.element_info.class_name}")
        print(f"{prefix}컨트롤 타입: {window.element_info.control_type}")
        print(f"{prefix}자동화ID: {window.element_info.automation_id}")
        print(f"{prefix}텍스트: {window.window_text()}")
        # MsgHtmlView가 포함된 컨트롤이면 하위 트리의 마지막(leaf) 텍스트만 출력
        if "MsgHtmlView" in window.element_info.class_name:
            print(f"{prefix}--- MsgHtmlView 하위 최하위 텍스트(최종값) ---")
            try:
                def get_last_leaf_text(ctrl):
                    try:
                        children = ctrl.children()
                        if not children:
                            return ctrl.window_text()
                        last_text = None
                        for c in children:
                            t = get_last_leaf_text(c)
                            if t:
                                last_text = t
                        return last_text
                    except Exception:
                        pass
                    return texts
                all_lines = collect_texts(window)
                merged_message = "\n".join(all_lines)
                print(f"{prefix}    [MERGED MESSAGE]\n{merged_message}")
            except Exception as e:
                print(f"{prefix}MsgHtmlView 하위 멀티라인 메시지 추출 실패: {e}")
    except Exception as e:
        print(f"{prefix}정보 출력 실패: {e}")
    # 자식 컨트롤 재귀 탐색
    try:
        children = window.children()
        for child in children:
            print_window_info(child, indent + 1, printed_ids)
    except Exception as e:
        print(f"{prefix}자식 컨트롤 탐색 실패: {e}")


def extract_message_text(msg_html_view):
    def collect_texts(ctrl):
        texts = []
        try:
            children = ctrl.children()
            if not children:
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
    return hashlib.sha256(message.encode("utf-8")).hexdigest()

def main():
    print(f"'{TARGET_PROCESS}' 쪽지창 디버깅(임시 전체 탐색)을 시작합니다...")
    printed_keys = set()
    while True:
        try:
            target_pids = [p.pid for p in psutil.process_iter(['name']) if p.info['name'] and p.info['name'].lower() == TARGET_PROCESS.lower()]
            for pid in target_pids:
                try:
                    app = Application(backend="uia").connect(process=pid)
                    windows = app.windows()
                    for win in windows:
                        # 모든 윈도우의 타이틀/클래스명 출력
                        print(f"윈도우 타이틀: {win.window_text()}, 클래스명: {win.element_info.class_name}")
                        try:
                            msg_html_views = [c for c in win.descendants() if "MsgHtmlView" in c.element_info.class_name]
                            for msg_html_view in msg_html_views:
                                message = extract_message_text(msg_html_view)
                                if message:
                                    key = get_message_key(message)
                                    if key not in printed_keys:
                                        print("\n===== 쪽지 본문 (임시 디버그) =====")
                                        print(f"[KEY] {key}")
                                        print(message)
                                        print("===============================\n")
                                        printed_keys.add(key)
                        except Exception as e:
                            print(f"MsgHtmlView 탐색 실패: {e}")
                except Exception as e:
                    print(f"PID {pid} 연결 실패: {e}")
            time.sleep(5)
        except KeyboardInterrupt:
            print("\n프로그램을 종료합니다.")
            sys.exit(0)

if __name__ == "__main__":
    main() 