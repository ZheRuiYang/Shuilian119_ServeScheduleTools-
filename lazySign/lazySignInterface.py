#! python3
# the interface of lazySign

import os, re, shelve, getpass

def schedulizer():
    '''開始值班登入排程或開始人員資料庫管理'''
    picPath = r'path\to\database\that\stroe\pictures' # for elegant, store .png as binary data stream
    dataPath = r'path\to\ID\database'#############################################################
    pyPath = r'path\to\python.exe'##############################################################
    pyw = r'path\to\pythonw.exe' ###############################################################
    scrPa = r'path\to\schedule\script.py' ###############################################################
    mList = []
    errList = {1:'資料庫裡沒有這個代號', 2:'看不懂這個指令'}
    
    print('┌──────────────┐ 　 パ..パ')
    print('│歡迎使用值班台登入排程小幫手│　（ ° ∀ ° ）')
    print('└──────────────┘＿（＿b /￣￣￣￣￣/＿')
    print('')
    print('請輸入指令：(「?」以查閱使用說明)')
    while True:
        _command = input()
        if re.match(r'(?|？)', _command):
            theHelp()
        elif re.match(r'(maintain|maintainning|M)', _command, re.IGNORECASE):
            dataMaintain()
        elif re.match('ok', _command) and not len(mList):
            _theChoosen = '; '.join(mList)
            os.system(f'{sciPa} {_theChoosen}')
            break
        else:
            if dataCheck(_command):
                print(f'{errList[int(dataCheck(_command))]}，請重新輸入：(「?」以查閱使用說明)')
            else:
                mList.append(_command)
                print('指令已記錄，目前已輸入的指令為：')
                for i in mList:
                    man = re.match(r'(.*)(,\d)+', i)
                    ti = [i.group(1)[1:] for i in re.finditer(r'(,\d)+', i)]
                    print(f'人員：{i.group(1)} 上值班台時間：', end='')
                    if len(ti) > 1:
                        for i in ti[:len(ti)]:
                            print(f'{i}時、', end='')
                        print(f'{ti[len(ti)]}時')
                    else:
                        print(f'{ti[0]}時')
                    print('')
                print('可以繼續輸入。(如欲開始排程請輸入「ok」；輸入「?」以查閱使用說明)')
                
def dataMaintain(dataPath):
    '''新增、修改、刪除或檢視人員資料'''
    with shelve.open(dataPath) as _id:
        while True:
            _which = input('輸入「a」新增人員；「d」刪除現有人員；「m」修改現有人員；「?」檢視現有人員；「q」回上一層：')
            if re.match(r'q', _which, re.IGNORECASE):
                backToSchedulizer()
            elif re.match(r'(?|？)', _which):
                showMember()
            elif re.match(r'a', _which, re.IGNORECASE): # Add new member
                while True:
                    mem = input('輸入想要的人員代號：(輸入「q」離開)')
                    if re.match(r'q', mem, re.IGNORECASE):
                        break
                    elif checkIDInDatabase(mem):
                        print('這個名稱已經有人使用。')
                    elif not checkIDInDatabase(mem):
                        print('基於保密原則，密碼輸入時螢幕上不會有回應。')
                        while True:
                            password = getpass.getpass(prompt='設定密碼：')
                            if not passwordCheck(password):
                                print('密碼不合身分證規則，請重新輸入。')
                            else:
                                _password = getpass.getpass(prompt='確認密碼：')
                                if not password == _password:
                                    print('設定密碼與確認密碼不符，請重新設定。')
                                else:
                                    while True:
                                        DTimes = input('輸入交接班作業的下拉式選單要按幾次下鍵才會指到你的名字：')
                                        if not isinstance(DTimes, int):
                                            print('次數必須要單純是個阿拉伯數字。')
                                        else:
                                            addMember([mem, password, DTimes])
                                            print(f'成功新增人員：{mem}')
                                            break
                                break
                        break
            elif re.match(r'd', _which, re.IGNORECASE): # delete existed member
                mem = input('輸入欲刪除的人員代號：(「?」檢視全部人員代號；「q」回上一層)')
                while True:
                    if re.match(r'q', mem, re.IGNORECASE):
                        break
                    elif re.match(r'(?|？)', _which):
                        showMember()
                    elif not checkIDInDatabase(mem):
                        print('該人員本來就不存在資料庫中。')
                    else:
                        while True:
                            print('基於保密原則，密碼輸入時螢幕上不會有回應。')
                            password = getpass.getpass(prompt='輸入密碼以確認刪除：')
                            if password == _id[mem] or password == _id[f'_{mem}']:
                                delMember(mem)
                                print(f'人員{mem}已成功刪除。')
                                break
                            else:
                                print('密碼錯誤！')
            elif re.match(r'm', _which, re.IGNORECASE): # modify existed member
                mem = input('輸入欲修改的人員代號：(「?」檢視全部人員代號；「q」回上一層)')
                while True:
                    if re.match(r'q', mem, re.IGNORECASE):
                        break
                    elif re.match(r'(?|？)', _which):
                        showMember()
                    elif not checkIDInDatabase(mem):
                        print('該人員不存在資料庫中。')
                    else:
                        print('基於保密原則，密碼輸入時螢幕上不會有回應。')
                        while True:
                            password = getpass.getpass(prompt='輸入密碼以確認修改：')
                            if password == _id[mem] or password == _id[f'_{mem}']:
                                password = getpass.getpass(prompt=f'確認修改人員{mem}，請設定新密碼：')
                                if not passwordCheck(password):
                                    print('密碼不合身分證規則，請重新輸入。')
                                else:
                                    _password = getpass.getpass(prompt='確認密碼：')
                                    if not password == _password:
                                        print('設定密碼與確認密碼不符，請重新設定。')
                                    else:
                                        while True:
                                        DTimes = input('輸入交接班作業的下拉式選單要按幾次下鍵才會指到你的名字：')
                                        if not isinstance(DTimes, int):
                                            print('次數必須要單純是個阿拉伯數字。')
                                        else:
                                            rewriteMember([mem, password, DTimes])
                                            print(f'人員{mem}已成功修改。')
                                            break
                                    break
                            else:
                                print('密碼錯誤！')
                    

    def addMember(imfo): # imfo = ['memberID', 'logInPassword', 'DownPressTimes']
        _id[imfo[0]] = imfo[1]
        _id[f'imfo[0]DownTimes'] = imfo[2]
        if imfo[1][0].isupper():
            imfo.append(f'{imfo[1][0].lower()}{imfo[1][1:]}')
        else:
            imfo.append(f'{imfo[1][0].upper()}{imfo[1][1:]}')
        _id[f'_{imfo[0]}'] = imfo[3]

    def delMember(ID):
        del _id[ID]
        del _id[f'_{ID}']
        del _id[f'{ID}DownTimes']

    def rewriteMember(imfo):
        addMember(imfo)
                
    def checkIDInDatabase(ID): # True－存在；False－不存在：
        if ID in _id:
            return True
        else:
            return False

    def showMember():
        for i in _id:
            print(i)
            
    def backToSchedulizer():
        break

    def passwordCheck(password):
        '''Check if password decently.Yes--True; No--False'''
        if re.match(r'[a-zA-Z]\d{9}', password):
            return True
        else:
            return False

def dataCheck(_string): # 檢查排程指令正確與否(return False/errorIndex)
    if re.match(r'(.*)(,\d)+', _string):
        _index = re.match(r'(.*)(,\d)+', _string).group(1)
        with shelve.open(dataPath) as _id:
            if not _index in id:
                return 1
            else:    
                return False
    else:
        return 2

def theHelp(): # 使用說明
    print('此程式可以幫忙自動登入值班台電腦。')
    print('輸入資料庫裡有的人員代號以及要上值班台時間可在該時間自動幫該人員登入。')
    print('新增、刪除及修改人員資料請輸入指令「maintain」以進入資料庫管理介面。')
    print('')
    print('●排程指令是「人員代號,時間,時間,時間......」。')
    print('  例如：代號ABC的人員預定要在08時與18時上值班台，請輸入：')
    print('  ABC,8,18')
    print('')
    print('PS.上值班台時間重複可能造成無法預期的結果，請小心確認所輸入的指令。')
