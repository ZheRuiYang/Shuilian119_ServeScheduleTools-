#! python3
# the interface of lazySign

import os, re, shelve, getpass

def schedulizer():
    '''開始值班登入排程或開始人員資料庫管理'''
    dataPath = r'path\to\ID\database'#############################################################
    pyPath = r'path\to\python.exe'##############################################################
    errList = {1:'資料庫裡沒有這個代號', 2:'看不懂這個指令'}
    mList = []
    print('┌──────────────┐ 　 パ..パ')
    print('│歡迎使用值班台登入排程小幫手│　（ ° ∀ ° ）')
    print('└──────────────┘＿（＿b /￣￣￣￣￣/＿')
    print('')
    print('請輸入指令：(「\help」以查閱使用說明；「\?」以檢視人員清單)')
    while True:
        _command = input()
        if re.match(r'(\\|/)(help|h)', _command, re.IGNORECASE):
            theHelp()
        elif re.match(r'(\\|/)(\?|？)', _command, re.IGNORECASE):
            with shelve.open(dataPath) as _id:
                showMember(_id)
                print('請輸入指令：(「\help」以查閱使用說明；「\?」以檢視人員清單)')
        elif re.match(r'(\\|/)(maintain|maintainning|M)', _command, re.IGNORECASE):
            dataMaintain(dataPath)
        elif re.match(r'(\\|/)(s|sche)', _command, re.IGNORECASE):
            if len(mList):
                showMList(mList)
                print('')
                print('可以繼續輸入。(輸入「\ok」開始行程；「\help」以查閱使用說明；「\?」以檢視人員清單；「\dele」以刪除已輸入的行程；「\sche」以檢視行程清單)')
            else:
                print('目前沒有行程。')
                print('可以繼續輸入。(輸入「\ok」開始行程；「\help」以查閱使用說明；「\?」以檢視人員清單；「\dele」以刪除已輸入的行程；「\sche」以檢視行程清單)')
        elif re.match(r'(\\|/)ok', _command, re.IGNORECASE):
            if len(mList):
                with shelve.open(dataPath) as _id:
                    for i in mList:
                        _co = i.split(',')[0]
                        if not _co in _id:
                            print(f'資料庫裡沒{_co}這個代號。')
                            print('可以繼續輸入。(輸入「\ok」開始行程；「\help」以查閱使用說明；「\?」以檢視人員清單；「\dele」以刪除已輸入的行程；「\sche」以檢視行程清單)')
                            print('')
                            break
                _theChoosen = ';'.join(mList)
                os.system(f'{scrPa} {_theChoosen}')
                break
            else:
                print('目前沒有行程。')
                print('可以繼續輸入。(輸入「\ok」開始行程；「\help」以查閱使用說明；「\?」以檢視人員清單；「\dele」以刪除已輸入的行程；「\sche」以檢視行程清單)')
        elif re.match(r'(\\|/)(d|dele)', _command, re.IGNORECASE):
            if len(mList):
                mList = modifyMList(mList)
                print('可以繼續輸入。(輸入「\ok」開始行程；「\help」以查閱使用說明；「\?」以檢視人員清單；「\dele」以刪除已輸入的行程；「\sche」以檢視行程清單)')
            else:
                print('目前沒有行程。')
                print('可以繼續輸入。(輸入「\ok」開始行程；「\help」以查閱使用說明；「\?」以檢視人員清單；「\dele」以刪除已輸入的行程；「\sche」以檢視行程清單)')
        else:
            _sep = _command.split(' ')
            if len(_sep) == 1 and len(_sep[0].split(',')) == 1:
                print('這是無效的指令。')
                print('------------')
                print('請輸入指令：(「\help」以查閱使用說明；「\?」以檢視人員清單)')
                continue
            for i in _sep:
                dc = dataCheck(i, dataPath)
                if dc:
                    print(f'{errList[int(dc)]}：{i}。請重新輸入：(「\help」以查閱使用說明；「\?」以檢視人員清單；「\sche」以檢視行程清單)')
                else:
                    mList.append(i)
            showMList(mList)
            print('')
            print('可以繼續輸入。(輸入「\ok」開始行程；「\help」以查閱使用說明；「\?」以檢視人員清單；「\dele」以刪除已輸入的行程；「\sche」以檢視行程清單)')

def showMList(aList):
    print('')
    print('●目前已輸入的指令為：')
    for i in aList:
        ma = i.split(',')
        print(f'{ma[0]}上值班台時間：', end='')
        for j in ma[1:]:
            if j != ma[len(ma)-1]:
                print(f'{j}時、', end='')
            else:
                print(f'{j}時')

def modifyMList(aList):
    print('<--刪除行程中人員-->')
    while True:
        bList = []
        _arg = input('輸入人員代號刪除以該行(輸入「\q」離開；「\sche」檢視目前行程)：')
        for i in aList:
            bList.append(i.split(',')[0])
        if re.match(r'(\\|/)q', _arg, re.IGNORECASE):
            print('離開刪除行程。')
            print('------------')
            break
        elif re.match(r'(\\|/)(s|sche)', _arg, re.IGNORECASE):
            showMList(aList)
            print('')
        elif _arg in bList:
            aList.remove(aList[bList.index(_arg)])
            print(f'成功刪除{_arg}的行程。')
            showMList(aList)
            print('')
        else:
            print('沒有這項行程。')
            print('')
    return aList
            
def dataMaintain(dataPath):
    '''新增、修改、刪除或檢視人員資料'''
    def addMember(imfo): # imfo = ['memberID', 'logInPassword', 'DownPressTimes']
        _id[imfo[0]] = imfo[1]
        _id[f'{imfo[0]}DownTimes'] = imfo[2]
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

    def passwordCheck(password):
        '''Check if password decently.Yes-->True; No-->False'''
        if re.match(r'[a-zA-Z]\d{9}', password):
            return True
        else:
            return False

    def passwordInput(modify=False):
        while True:
            if modify:
                password = getpass.getpass(prompt=f'確認修改人員{modify}，請設定新密碼(「\q」回上一層)：')
            else:
                password = getpass.getpass(prompt='設定密碼(「\q」回上一層)：')
            if re.match(r'(\\|/)q', password, re.IGNORECASE):
                return password
            elif not passwordCheck(password):
                print('密碼不符合身分證規則。')
            else:
                _password = getpass.getpass(prompt='確認密碼：')
                if not password == _password:
                    print('設定密碼與確認密碼不符，請重新設定。')
                else:
                    return password

    def downTimesInput():
        while True:
            DTimes = int(input('輸入交接班作業的下拉式選單要按幾次下鍵才會指到你的名字：'))
            if not isinstance(DTimes, int):
                print('次數必須要單純是個阿拉伯數字。')
            else:
                return DTimes
        
    with shelve.open(dataPath) as _id:
        while True:
            print('【資料庫維護】')
            _which = input('輸入「\add」新增人員；「\dele」刪除現有人員；「\modi」修改現有人員；「\?」檢視現有人員；「\q」回上一層：')
            if re.match(r'(\\|/)q', _which, re.IGNORECASE):
                print('=============================================')
                print('請輸入指令：(「\help」以查閱使用說明；「\?」以檢視人員清單)')
                break
            elif re.match(r'(\\|/)(\?|？)', _which):
                showMember(_id)
            elif re.match(r'(\\|/)(a|add)', _which, re.IGNORECASE): # Add new member
                while True:
                    print('')
                    print('<---新增人員--->')
                    mem = input('輸入想要的人員代號(輸入「\q」回上一層)：')
                    if re.match(r'(\\|/)q', mem, re.IGNORECASE):
                        print('離開新增人員。')
                        print('------------')
                        print('')
                        break
                    elif re.match(r'(\\|/)', mem):
                        sla = re.match(r"(\\|/)", mem).group()
                        print(f'名稱不可以「{sla}」開頭。')
                    elif checkIDInDatabase(mem):
                        print('這個名稱已經有人使用。')
                    elif not checkIDInDatabase(mem):
                        print('※基於保密原則，密碼輸入時螢幕上不會有回應。')
                        password = passwordInput()
                        if re.match(r'(\\|/)q', password, re.IGNORECASE):
                            continue
                        DTimes = downTimesInput()
                        addMember([mem, password, DTimes])
                        print(f'成功新增人員：{mem}')
                        print('------------')
                        print('')
                        break
            elif re.match(r'(\\|/)(d|dele)', _which, re.IGNORECASE): # delete existed member
                while True:
                    print('')
                    print('<---刪除人員--->')
                    mem = input('輸入欲刪除的人員代號(「\?」檢視全部人員代號；「\q」回上一層)：')
                    if re.match(r'(\\|/)q', mem, re.IGNORECASE):
                        print('離開刪除人員。')
                        print('------------')
                        print('')
                        break
                    elif re.match(r'(\\|/)(\?|？)', mem):
                        showMember(_id)
                    elif not checkIDInDatabase(mem):
                        print('該人員不存在資料庫中。')
                        print('')
                    else:
                        print('※基於保密原則，密碼輸入時螢幕上不會有回應。')
                        while True:
                            password = getpass.getpass(prompt='輸入該人員的密碼以確認刪除：')
                            if password == _id[mem] or password == _id[f'_{mem}']:
                                delMember(mem)
                                print(f'已成功刪除人員{mem}。')
                                break
                            else:
                                print('密碼錯誤！')
            elif re.match(r'(\\|/)(m|modi)', _which, re.IGNORECASE): # modify existed member
                while True:
                    print('')
                    print('<---修改人員資料--->')
                    mem = input('欲修改的人員代號(「\?」檢視全部人員代號；「\q」回上一層)：')
                    if re.match(r'(\\|/)q', mem, re.IGNORECASE):
                        print('離開修改人員。')
                        print('------------')
                        print('')
                        break
                    elif re.match(r'(\\|/)(\?|？)', mem):
                        showMember(_id)
                    elif not checkIDInDatabase(mem):
                        print('該人員不存在資料庫中。')
                        print('')
                    else:
                        print('※基於保密原則，密碼輸入時螢幕上不會有回應。')
                        password = getpass.getpass(prompt='輸入該人員的密碼以確認修改：')
                        if password == _id[mem] or password == _id[f'_{mem}']:
                            password = passwordInput(modify=f'{mem}')
                            if re.match(r'(\\|/)q', password, re.IGNORECASE):
                                continue
                            MDTCheck = input('是否要一起修改下鍵次數？(「\y」確認；其餘取消)')
                            if re.match(r'(\\|/)y', MDTCheck, re.IGNORECASE):
                                DTimes = downTimesInput()
                            else:
                                DTimes = _id[f'{mem}DownTimes']
                            rewriteMember([mem, password, DTimes])
                            print(f'已成功修改{mem}的人員資料。')
                            print('------------')
                            print('')
                            break
                        else:
                            print('密碼錯誤！')
            else:
                print('這是無效的指令。')
                print('------------')
                print('')
                    
def dataCheck(_string, dataPath): # return: 0-->OK, 1-->not in DB, 2-->can't understand)
    _st = _string.split(',')
    for i in _st[1:]:
        if not re.match(r'\d{1,2}', i):
            return 2
    with shelve.open(dataPath) as _id:
        if _st[0] in _id:
            return 0
        else:
            return 1
        
def showMember(_id):
    print('')
    print('●人員清單：')
    if len(_id):
        for i in _id:
            if re.search(r'(_|DownTimes)', i):
                continue
            else:
                print(i)
    else:
        print('資料庫現在是空著的。')
    print('')

def theHelp():
    '''使用說明'''
    print('–––––––––––––––––––––––––––––––––––')
    print('此程式可以幫忙自動登入值班台電腦。')
    print('輸入資料庫裡有的人員代號以及要上值班台時間可在該時間自動幫該人員登入。')
    print('新增、刪除及修改人員資料請輸入指令「\maintain」以進入資料庫管理介面。')
    print('')
    print('●排程指令是「人員代號,時間,時間,時間......」。')
    print('  例如：代號ABC的人員預定要在08時與18時上值班台，請輸入：')
    print('  ABC,8,18')
    print('')
    print('注意！上值班台時間重複可能造成無法預期的結果，請小心確認所輸入的指令。')
    print('–––––––––––––––––––––––––––––––––––')
    print('可以繼續輸入。(輸入「\ok」開始行程；「\help」以查閱使用說明；「\?」以檢視人員清單；「\dele」以刪除已輸入的行程；「\sche」以檢視行程清單)')

if __name__ == '__main__':
    schedulizer()
