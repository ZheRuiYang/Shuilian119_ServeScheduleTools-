#! python3
# auto sign in serve table

import pyautogui, time, sched, os, shelve, sys

def interface(window=True, _schedule=None):
    try:
        _schedule = sys.argv[1]
        for i in _schedule.split(';'):
            preOrder(i)
    except:
        a = []
        while True:
            if window: # open a window to collecting imformations
                a.append(input('輸入人員代號及上值班台時間: (e.g. ○○○,10,20) '))
                _pause = input('輸入完成？ (輸入「y」以檢查輸入的資料；其餘任意鍵以繼續輸入。)')
                if _pause == 'y':
                    print('目前已輸入的資料：')
                    for i in a:
                        print(i)
                    print('')
                    _excu = input('確認無誤後輸入「go」以開始；其餘任何鍵以繼續輸入。') # delete?
                    if _excu == 'go':
                        window = False
                    else:
                        continue
            else: # run the main process without window
                f = ';'.join(a)
                pyw = r'path\to\pythonw.exe' ###############################################################
                scrpa = r'path\to\script.py' ###############################################################
                os.system('pause >nul | echo 排程已記錄，按下任意鍵以轉入背景運行...')
                os.system(f'{pyw} {scrpa} {f}')
                break

def preOrder(t):
    sch = sched.scheduler(time.time, time.sleep)
    for i in t.split(',')[1:]:
        sch.enterabs(epochTime(i), 1, signIn, argument=(t.split(',')[0],))
    sch.enterabs(epochTime(22, end=True), 1, os._exit)
    sch.run()
    
def epochTime(h, end=False):
    if end:
        dayString = time.strftime('%Y%m%d')
        return time.mktime((int(dayString[:4]), int(dayString[4:6]), int(dayString[6:]), int(h)-1, 59, 18, 0, 0, -1))
    else:
        dayString = time.strftime('%Y%m%d')
        return time.mktime((int(dayString[:4]), int(dayString[4:6]), int(dayString[6:]), int(h), 30, 18, 0, 0, -1))
    
def signIn(name):
    # locate app
    picPath = os.path.join(os.environ['USERPROFILE'], 'AppData\Local\lazySign')
    counter = 0
    while True:
        if not pyautogui.locateCenterOnScreen(os.path.join(picPath, '交接班管理作業.png')):
            counter +=1
            pyautogui.keyDown('alt')
            pyautogui.press('\t', presses=counter)
            pyautogui.keyUp('alt')
        else:
            break
    # key in
    while True:
        pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '交接班管理作業.png')), interval=0.3)
        if pyautogui.locateCenterOnScreen(os.path.join(picPath, '交接班作業.png')):
            pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '交接班作業.png')), interval=0.1)
            break
        else:
            continue
    downPressTimes = {'ZRY': 7}
    pyautogui.press('down', presses=downPressTimes[f'{name}'], interval=0.1)
    pyautogui.typewrite('\t')
    with shelve.open(picPath + '\PID') as id:
        pyautogui.typewrite(id[f'{name}'], interval=0.1)
    pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '確定.png')), interval=1)
    time.sleep(5)
    # check whether sign-in success
    while True:
        if pyautogui.locateCenterOnScreen(os.path.join(picPath, '密碼錯誤.png')):
            pyautogui.hotkey('ctrl', 'a', interval=0.1)
            pyautogui.press('backspace', interval=0.5)
            with shelve.open(picPath + '\PID') as id:
                pyautogui.typewrite(id[f'_{name}'], interval=0.1)
            pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '確定.png')), interval=1)
        else:
            break
        
if __name__ == '__main__':
    interface()
