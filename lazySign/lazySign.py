#! python3
# auto sign in serve table

import pyautogui, time, sched, os, shelve, sys, re

def interface(window=True, _schedule=None):
    picPath = os.path.join(os.environ['USERPROFILE'], 'AppData\Local\lazySign')
    pyPath = r'path\to\python.exe'##############################################################
    pyw = r'path\to\pythonw.exe' ###############################################################
    scrPa = r'path\to\this\script.py' ###############################################################
    try:
        _schedule = sys.argv[1]
        for i in _schedule.split(';'):
            preOrder(i)
    except:
        _dataBox = []
        print('┌──────────────┐ 　 パ..パ')
        print('│歡迎使用值班台登入排程小幫手│　（ ° ∀ ° ）')
        print('└──────────────┘＿（＿b /￣￣￣￣￣/＿')
        print('')
        while True:
            if window: # open a window to collecting imformations
                _text = input('輸入人員代號及上值班台時間；直接按enter以跳過；輸入「member」以進入新增/刪除人員模式： (e.g. ○○○,10,20) ')
                if _text == 'member':
                    _switch = input('輸入「new」以新增人員資料；輸入「del」以刪除人員資料；輸入其餘任何鍵以回到上一層。')
                    if _switch == 'new':
                        while True:
                            _newmember = []
                            _newmember.append(input('輸入人員代號：(如不想新增，直接enter以退回上一層)：'))
                            if _newmember[0] == '':
                                break
                            _newmember.append(input('輸入交接班作業的下拉選單要按幾次下鍵才能正確指定：'))
                            _newmember.append(input('輸入登入密碼：(注意：此程式會永久記錄此密碼直到手動刪除！)'))
                            if _newmember[2][0].upper():
                                _newmember.append(f'{_newmember[2][0].lower()}{_newmember[2][1:]}')
                            else:
                                _newmember.append(f'{_newmember[2][0].upper()}{_newmember[2][1:]}')
                            with shelve.open(picPath + '\PID') as id:
                                if not isinstance(_newmember[1], int) or not re.match(r'^[A-Za-z](\d){9}$', _newmember[2]):
                                    if not isinstance(_newmember[1], int):
                                        print('交接班作業的下拉選單部分必須要是數字，請全部重新輸入。')
                                    else:
                                        print(f'您輸入的登入密碼「{_newmember[2]}」不符合格式，請全部重新輸入。')
                                elif _newmember[0] in id:
                                    print(f'您所設定的人員代號為{_newmember[0]}，資料庫內部已有這筆資料。')
                                else:
                                    print(f'您所設定的人員代號為{_newmember[0]}；交接班作業下拉選單的下鍵指定次數為{_newmember[1]}；登入密碼為{_newmember[2]}。')
                                    if _newmember[0] in id:
                                        print('')
                                        print('※※ 資料庫內部已有這筆資料 ※※')
                                        print('')
                                        _parad = input('輸入「ok」以覆蓋資料庫內部資料；其餘任何鍵以重新輸入資料。')
                                        if not re.match(r'(ok|OK|Ok|oK)', _parad):
                                            continue
                                    _memberDataCheck = input('確認無誤後輸入「ok」將資料存入電腦；輸入其餘任何鍵重新輸入：')                                
                                    if re.match(r'(ok|OK|Ok|oK|okey|Okey|okay|Okay|okey-doke)', _memberDataCheck):
                                        id[f'{_newmember[0]}'] = _newmember[2]
                                        id [f'_{_newmember[0]}'] = _newmember[3]
                                        id[f'{_newmember[0]}_KeyPressTimes'] = _newmember[1]
                                        os.system('pause >nul | echo 資料儲存完成，按下任何鍵以回上一層。')
                                        break
                                    else:
                                        continue
                    elif _switch == 'del':
                        while True:
                            _useless = input('輸入要刪除的人員代號以刪除該筆資料；輸入「?」以閱覽所有的人員代號；輸入「q」離開刪除模式：')
                            if _useless == '?':
                                with shelve.open(picPath + '\PID') as id:
                                    for i in id.keys():
                                        print(i)
                            elif _useless == 'q':
                                os.system(f'{pyPath} {scrPa}')
                                os._exit()
                            else:
                                with shelve.open(picPath + '\PID') as id:
                                    if not _useless in id:
                                        print('資料庫中沒有這個人員代號。')
                                    else:
                                        delCheck = input(f'確認刪除{_useless}？ (輸入對應的登入密碼以確認刪除；輸入「n」取消。)')  
                                        if delCheck == id[_useless]:
                                            del id[_useless]
                                            del id[f'{_useless}_KeyPressTimes']
                                            os.system('pause >nul | echo 成功刪除{_useless}！按下任何鍵以繼續。')
                                        elif delCheck == 'n':
                                            continue
                                        else:
                                            print('密碼錯誤，請再次確認。')
                    else:
                        continue
                elif re.match(r'(.*)(,(\d|\d\d))+', _text):
                    _dataBox.append(_text)
                else:
                    print(f'你輸入的字串是「{_text}」。該字串格式有誤，請重新輸入。')
                    print('')
                    continue
                _pause = input('輸入完成？ (輸入「y」以檢查輸入的資料；輸入其餘任意鍵以繼續輸入。)')
                if _pause == 'y':
                    while True:
                        print('目前已輸入的資料：')
                        for i in _dataBox:
                            print(i)
                        print('')
                        _excu = input('確認無誤後輸入「go」以開始；輸入「d」以刪除已輸入的資料；輸入其餘任何鍵以繼續輸入。')
                        if _excu == 'go':
                            window = False
                            break
                        elif _excu == 'd':
                            _d = input('輸入人員代號以刪除該行資料：(e.g. 欲刪除「abc,08,18」這行請輸入「abc」) ')
                            for i in _dataBox:
                                try:
                                    del _dataBox[re.match(f'{_d}.*', i).group()]
                                except:
                                    continue
                        else:
                            break
            else: # run the main process without window
                _dataString = ';'.join(_dataBox)
                os.system('pause >nul | echo 排程已記錄，按下任意鍵以轉入背景運行...')
                os.system(f'{pyw} {scrPa} {_dataString}')
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
    picPath = os.path.join(os.environ['USERPROFILE'], 'AppData\Local\lazySign')
    # locate app
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
    with shelve.open(picPath + '\PID') as id:
        pyautogui.press('down', presses=int(id[f'{name}_KeyPressTimes']), interval=0.2)
        pyautogui.typewrite('\t')
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
