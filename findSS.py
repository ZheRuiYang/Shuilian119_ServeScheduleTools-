def findSS():
    import re, os

    print('這裡可以查詢Dropbox裡特定警消、役男人數組合的勤務表。')
    print('若要搜尋3警消2役男的組合，請輸入32。')
    print('請輸入想查找的組合(警消2~4、役男0~5)：', end='')
    kind = input()
    
    para = list(kind)[0] + '+' + list(kind)[1]
    patt = '(\d){7}(\(' + re.escape(para) + '\))(.*)$'
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
    os.system('pause')

if __name__ == '__main__':
    findSS()
