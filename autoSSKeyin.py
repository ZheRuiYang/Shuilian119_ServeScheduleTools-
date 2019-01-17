#! python3
# Auto key-in serve schedule to online system

import re, pyautogui, os, docx, time, pyperclip

def main():
    print('┌───────────┐   　.  ^＿^ ∩* ☆')
    print('│勤務表輸入小幫手已開啟│ * ☆ (˙▽˙)/ .')
    print('└───────────┘　 . ⊂　　 ノ* ☆')
    print('')
    while True:
        print('請輸入對應勤務表日期： (範例－西元2018年5月14日為「1070514」)')
        dat = input()
        tt = tuple([int(dat[:3]) + 1911, int(dat[3:5]), int(dat[5:]),0,0,0,0,0,-1])
        try:
            time.asctime(tt)
            break
        except:
            print('日期格式不對，請重新輸入')
            continue
    for root, dirs, files in os.walk(os.path.join(os.environ['USERPROFILE'], 'Dropbox')):
        for File in files:
            mat = re.match(dat + r'\(\d\+\d\)(一|二|三|四|五|六|日).*\.docx$', File)
            if mat:
                if re.search('deleted', mat.group()):
                    continue
                else:
                    case = mat
                    filePath = root
                    break
    # read SS
    data = [[], [], []] # [[tbl#3][tbl#4][tbl#6]]
    try:
        tbl = docx.Document(os.path.join(filePath, case.group())).tables
    except UnboundLocalError:
        print(f'Dropbox裡面沒有{dat[:3]}年{dat[3:5]}月{dat[5:]}日的勤務表')
        os.system('pause >nul')
    for row in tbl[2].rows:
        for cell in row.cells:
            data[0].append(cell.text.replace('\xa0', '\t'))
    for row in tbl[3].rows:
        for cell in row.cells:
            data[1].append(cell.text.replace('\xa0', '\t'))
    for row in tbl[5].rows:
        for cell in row.cells:
            data[2].append(cell.text.replace('\xa0', '\t'))
    useless = [i for i in data[0][:9]]
    for i in useless:
        data[0].remove(i)
    useless = [data[0][i] for i in range(len(data[0])) if re.search(r'\d{1,2}~\d{1,2}', data[0][i])]
    for i in useless:
        data[0].remove(i)
    useless = [data[2][i] for i in range(len(data[2])) if re.match('出動梯次', data[2][i]) or re.match('附記', data[2][i])]
    for i in useless:
        data[2].remove(i)
    box = data[2][1].split('\n')
    del data[2][1]
    header = {1: '一、', 2:'二、', 3:'三、', 4: '四、', 5: '五、', 6: '六、', 7: '七、', 8: '八、', 9: '九、', 10: '十、'}
    for i in range(len(box)):
        data[2].append(header[i+1] + box[i])
    # key-in
    print('您選擇的日期為：民國' + case.group()[:3] + '年' + case.group()[3:5] + '月' + case.group()[5:7] + '日')
    print('請將指標停在最左上角(「8~9」, 「值班」)那格')
    print('確認無誤後輸入任何鍵以開始自動程序前的5秒倒數')
    os.system('pause')
    for i in range(5, 0, -1):
        print('自動輸入將於' + str(i) + '秒後開始...')
        time.sleep(1)                
    for i in range(len(data[0])):
        pyautogui.typewrite(data[0][i])
        pyautogui.typewrite('\t')
    for i in range(6):
        pyautogui.typewrite(data[1][i*2+1])
        pyautogui.typewrite('\t')
    pyautogui.typewrite('\t')
    pyperclip.copy(data[2][0])
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.typewrite('\t')
    for i in data[2][1:]:
        pyperclip.copy(i)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.typewrite(['enter'])
                    
if __name__ == '__main__':
    main()
