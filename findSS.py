#! python3
#

import re, os, docx, subprocess

def interface():
    print('┌──────┐')
    print('│勤務表搜尋器│')
    print('└──────┘')
    # argument parser
    while True:
        _arg = input('輸入搜尋條件：(輸入 ? 以閱讀操作說明)')
        if re.match(r'(\?|？)', _arg):
            instructionManual()
        else:
            break
    print('正在搜尋中...')
    result = searchManager(_arg)
    if len(result) != 0:
        print('')
        print('搜尋結果如下：')
    for i in range(len(result)):
        print(f'{i+1}.', result[i][1])
    print('搜尋完畢！')
    print('')
    if len(result) == 0:
        print('沒有這種勤務表。')
        os.system('pause >unl | echo 按下任何鍵以關閉')
    else:
        while True:
            _index = input('輸入序號以開啟對應勤務表的所在位置：')
            try:
                _index = int(_index)-1
            except ValueError:
                Print('序號要是個阿拉伯數字。')
                continue
            try:
                _path = os.path.join(result[_index][0], result[_index][1])
                _path = os.path.normpath(_path)
            except IndexError:
                print('請重新確認該序號是否存在。')
                print('')
                continue
            print(f'開啟{_index+1}. {result[_index][1]}')
            print('')
            subprocess.Popen([r'explorer', r'/select,', _path])

def searchManager(requirement):
    print('分析搜尋條件...', end='')
    # parsing argument
    _filterCounter = 0
    if re.search(r'"(.*)"', requirement):
        root = re.search(r'"(.*)"', requirement).group(1)
    else:
        root = None
    if re.search(r'PS(:|：)(\w+)', requirement, re.IGNORECASE):
        ps = re.search(r'PS(:|：)(\w+)', requirement, re.IGNORECASE).group(2)
        _filterCounter += 1
    else:
        ps = None
    if re.search(r'(CONTENT|C)(:|：)(\w+)', requirement, re.IGNORECASE):
        content = re.search(r'(CONTENT|C)(:|：)(\w+)', requirement, re.IGNORECASE).group(3)
        _filterCounter += 1
    else:
        content = None
    if re.search(r'(\d)\+(\d)', requirement):
        combo = re.search(r'(\d)\+(\d)', requirement).group()
        _filterCounter += 1
    else:
        combo = None
    if re.search(r'(MEMBER|M)(:|：)(\d)(,\d)*', requirement, re.IGNORECASE):
        member = re.search(r'(MEMBER|M)(:|：)(\d)(,\d)*', requirement, re.IGNORECASE).group()
        _filterCounter += 1
    else:
        member = None
    print('完成！')
    # record filenames (and their path) under root
    if root:
        cand = theWalker(root)
    else:
        cand = theWalker()
    print('檢查檔案中...', end='')
    _ok = []
    # analysis by filenames
    for i in range(len(cand)):
        if ps:
            if byNamePS(cand[i][1], ps):
                _ok.append(cand[i])
        if combo:
            if byMemberCombo(cand[i][1], combo):
                _ok.append(cand[i])
        if member:
            if byMember(cand[i], member):
                _ok.append(cand[i])
        if content:
            if byContent(cand[i], content):
                _ok.append(cand[i])
    print('完成！')
    print('收集結果中...', end='')
    _result = []
    for i in range(len(_ok)):
        _counter = 0
        _sub = _ok.pop(0)
        for j in _ok:
            if _sub == j:
                _counter += 1
        if _counter == _filterCounter-1:
            _result.append(_sub)
    print('完成！')
    return _result

def byNamePS(target, ps):
    patt = '^(\d){7}\(\d\+\d\)\w.*' + re.escape(ps) + '.*$'
    return re.search(patt, target)

def byMemberCombo(target, combo): # (\d)+(\d)
    patt = '^(\d){7}(\(' + re.escape(combo) + '\))(.*)$'
    return re.search(patt, target)

def byMember(target, member):
    member = member.split(',')
    member[0] = member[0][len(member[0])-1]
    leave = []
    tbl = docx.Document(os.path.join(target[0], target[1])).tables
    try:
        tbl = tbl[3]
        for row in tbl.rows:
            for cell in row.cells:
                text = cell.text.replace('\xa0', '\t')
                text = text.split(',')
                for i in text:
                    leave.append(i)
        for i in member:
            for j in leave:
                if j == i:
                    return None
    except IndexError:
        return None
    return 1

def byContent(target, content):
    #print(content)
    textPile = []
    tbls = docx.Document(os.path.join(target[0], target[1])).tables
    for tbl in tbls[3:]:
        for row in tbl.rows:
            for cell in row.cells:
                text = cell.text.replace('\xa0', '\t')
                if re.search(content, text):
                    return 1
    
def theWalker(start=os.path.join(os.environ['USERPROFILE'], 'Dropbox')):
    print('收錄候選檔案中...', end='')
    target = []
    for root, dirs, files in os.walk(start):
        for file in files:
            file = re.match(r'^\d{7}.*(.docx)$', file)
            if file:
                file = file.group()
                if re.search(r'deleted', file):
                    continue
                target.append([root, file])
            else:
                continue
    print('完成！')
    return target

def instructionManual():
    print('')
    print('使用說明：')
    print('● 搜尋特定的警消+役男組合請輸入數字組合。')
    print('● 搜尋勤務表名稱的備註請在文字前加上「PS:」。')
    print('● 搜尋勤務表內容請在文字前加上「CONTENT:」。')
    print('● 搜尋特定番號上班的勤務表請在番號數字前加上「MEMBER:」，兩位以上用「,」隔開。')
    print('● 搜尋指定資料夾請把資料夾路徑用英文雙引號(")括起來。(預設是Dropbox)')
    print('')
    print('多重條件搜尋請把每個條件用空白分開。')
    print('舉例：搜尋位於C:\底下、警消與役男組合為3+2、1,4,7番同時有上班、勤務表名稱有「慈恩」兩字、內容含有「體能訓練」四字的勤務表要輸入：')
    print('"C:\" 3+2 MEMBER:1,4,7 PS:慈恩 CONTENT:體能訓練')
    print('')
    print('搜尋條件順序不拘。')
    print('注意：搜尋內容與搜尋特定番號會多花上許多時間。')
    print('')

if __name__ == '__main__':
    interface()
