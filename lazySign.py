#! python3
# auto sign in serve table

import pyautogui, time, sched, os, shelve

def preOrder(t):
    sch = sched.scheduler(time.time, time.sleep)
    for i in t.split(','):
        sch.enterabs(epochTime(i), 1, signIn)
    sch.enterabs(epochTime(22), 1, os._exit)
    sch.run()
    
def epochTime(h):
    dayString = time.strftime('%Y%m%d')
    return time.mktime((int(dayString[:4]), int(dayString[4:6]), int(dayString[6:]), int(h)-1, 59, 18, 0, 0, -1))
    
def signIn():
    # locate app
    counter = 0
    while True:
        if not pyautogui.locateCenterOnScreen(os.path.join(picPath, '交班管理作業.png')):
            counter +=1
            pyautogui.keyDown('alt')
            pyautogui.press('\t', presses=counter)
        else:
            pyautogui.keyUp('alt')
            break
    # key in
    picPath = os.environ['USERPROFILE'] + r'\AppData\Local\lazySign'
    pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '交班管理作業.png')), interval=0.3)
    pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '交班作業.png')), interval=0.1)
    pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '請選擇.png')), interval=0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, 'ZRY.png')), interval=1) ###################################################
    pyautogui.typewrite('\t')
    with shelve.open(picPath + '\PID') as id:
        pyautogui.typewrite(id['ZRY'], interval=0.1) ###################################################
    pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '確定.png')), interval=1)
    # check whether sign-in success
    while True:
        if pyautogui.locateCenterOnScreen(os.path.join(picPath, '密碼錯誤.png')):
            pyautogui.hotkey('ctrl', 'a', interval=0.1)
            pyautogui.press('backspace', interval=0.5)
            with shelve.open(picPath + '\PID') as id:
                pyautogui.typewrite(id['_ZRY'], interval=0.1) ###################################################
            pyautogui.click(pyautogui.locateCenterOnScreen(os.path.join(picPath, '確定.png')), interval=1)
        else:
            break
        
if __name__ == '__main__':
    preOrder(input())
