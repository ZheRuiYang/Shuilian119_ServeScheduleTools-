def findSS():
    import re, os

    while True:
        print('這裡可以查詢Dropbox裡特定警消、役男人數組合的勤務表。(只比對檔案名稱)')
        print('')
        print('● 若要搜尋三警消二役男的組合，請輸入32。')
        print('● 搜尋全範圍則輸入"?"，例如三役男與任何警消組合請輸入?3。')
        print('● 搜尋備註請直接輸入關鍵字，例如輸入慈恩。')
        print('')
        kind = input('請輸入想查找的組合(警消2~4、役男0~5)：')
        if re.match(r'\d\d', kind) or re.match(r'\?\d', kind) or re.match(r'\d\?', kind) or re.match(r'\w+', kind):
            break
        else:
            print('你的輸入不符合要求，請重新輸入。')
            continue

    if re.match(r'\d\d', kind):
        para = re.escape(f'{kind[0]}+{kind[1]}')
        patt = '^(\d){7}(\(' + para + '\))(.*)$'
    elif re.match(r'\?\d', kind):
        para = re.escape(f'+{kind[1]}')
        patt = '^(\d){7}(\(\d' + para + '\))(.*)$'
    elif re.match(r'\d\?', kind):
        para = re.escape(f'{kind[0]}+')
        patt = '^(\d){7}(\(' + para + '\d\))(.*)$'
    elif re.match(r'\w+', kind):
        para = re.escape(re.match(r'\w+', kind).group())
        patt = '^(\d){7}\(\d\+\d\)\w.*' + para + '.*$'
        
    filenameRegex = re.compile(patt)
    res = []
    
    dropbox = os.path.join(os.environ['USERPROFILE'], 'Dropbox')
    for folderName, subfolders, filenames in os.walk(dropbox):
        for filename in filenames:
            gotOne = filenameRegex.search(filename)
            if gotOne == None:
                continue
            else:
                res.append(gotOne.group())

    if len(res) == 0:
        print('')
        print('沒有那種勤務表。')
    else:
        for i in res:
            print(i)
        print('搜尋完畢。')
    print('')
    os.system('pause >nul | echo 按下任意鍵以關閉視窗...')

if __name__ == '__main__':
    findSS()
