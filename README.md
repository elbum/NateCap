# Nateonbiz Capture

## dependacy
python 3.11
pywinauto-0.6.9
psutil-7.0.0

## download packages
pip download pywinauto --only-binary=:all: \                
     --platform win_amd64 \
     --python-version 311 \
     --abi cp311 \
     --implementation cp \
     -d .
    
## install packages
pip install -v --no-index --find-links . pywin32

## run
py natecap.py

