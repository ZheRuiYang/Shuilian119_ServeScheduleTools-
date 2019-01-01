#! python3
# Batch make serve schdule for a whole month.  

import os, zipfile, re, datetime, shutil, codecs
import xml.etree.ElementTree as ET

def xmlSetup(tree):
    atts = {"mc:Ignorable": "w14 w15 w16se wp14",
            "xmlns:wps": r"http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
            "xmlns:wne": r"http://schemas.microsoft.com/office/word/2006/wordml",
            "xmlns:wpi": r"http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
            "xmlns:wpg": r"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "xmlns:w16se": r"http://schemas.microsoft.com/office/word/2015/wordml/symex",
            "xmlns:w15": r"http://schemas.microsoft.com/office/word/2012/wordml",
            "xmlns:w14": r"http://schemas.microsoft.com/office/word/2010/wordml",
            "xmlns:w10": r"urn:schemas-microsoft-com:office:word",
            "xmlns:wp": r"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "xmlns:wp14": r"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "xmlns:v": r"urn:schemas-microsoft-com:vml",
            "xmlns:m": r"http://schemas.openxmlformats.org/officeDocument/2006/math",
            "xmlns:r": r"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "xmlns:o": r"urn:schemas-microsoft-com:office:office",
            "xmlns:mc": r"http://schemas.openxmlformats.org/markup-compatibility/2006",
            "xmlns:cx": r"http://schemas.microsoft.com/office/drawing/2014/chartex",
            "xmlns:wpc": r"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"}
    
    root = ET.Element(r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}document", atts)
    body = tree.find(r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")
    root.append(body)
    treeString = ET.tostring(root, encoding="unicode").replace('\xa0', '\t')
    treeString = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + treeString.replace('ns0', 'w')
    treeString = codecs.encode(treeString)
    with open("document.xml", "wb") as xml:
        xml.write(treeString)

def tbl34Content(tree, row, column, value, which): # table 3 and table 4
    w = r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    
    which = which - 1
    cPara = tree.find(w + 'body').findall(w + 'tbl')[which].findall(w + 'tr')[int(row)].findall(w + 'tc')[int(column)].find(w + 'p')
    wr = wrElement(value)
    cPara.append(wr)
    wpPr = cPara.find(w + 'pPr')
    wrPr = ET.SubElement(wr, w + 'rPr')
    ET.SubElement(wrPr, w + 'rFonts', attrib={w + "hint": "eastAsia"})
    ET.SubElement(wrPr, w + 'sz', attrib={w + "val": "20"})
    ET.SubElement(wrPr, w + 'szCs', attrib={w + "val": "20"})
    return tree

def drwningProof(tree):
    w = r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    
    drownPr = tree.find(w + 'body').findall(w + 'tbl')[2].findall(w + 'tr')[0].findall(w + 'tc')[4].find(w + 'p')
    drownPr.remove(drownPr.find(w + 'r'))
    drownPr.append(wrElement('防溺宣導'))
    return tree

def tbl2Content(tree, date):
    w = r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    d = re.compile(r"(\d\d\d)(\d\d)(\d\d)")
    mo = d.search(date)
    date = mo.group(1) + "/" + mo.group(2) + "/" + mo.group(3)

    datePara = tree.find(w + 'body').findall(w + 'tbl')[1].findall(w + 'tr')[2].findall(w + 'tc')[0].find(w + 'p').find(w + 'r')
    ET.SubElement(datePara, w + 't').text = date
    return tree

def tbl6Content(tree, PS1, PS2): # PS1=出動梯次(string)、PS2=附記(list)
    w = r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    wp1 = tree.find(w + 'body').findall(w + "tbl")[5].findall(w + "tr")[0].findall(w + "tc")[1].find(w + "p")
    wp1.append(wrElement(PS1))
    
    wtc2 = tree.find(w + 'body').findall(w + "tbl")[5].findall(w + "tr")[1].findall(w + "tc")[1]
    wtc2.clear()
    wtc2Pr = ET.Element(w + "tcPr")
    ET.SubElement(wtc2Pr, w + "tcW", attrib={w + "w": "0", w + "type": "auto"})
    wtc2B = ET.SubElement(wtc2Pr, w + "tcBoarders")
    atts = {w + "val": "outset", w + "space": "0", w + "color": "auto", w + "sz": "6"}
    ET.SubElement(wtc2B, w + "top", atts)
    ET.SubElement(wtc2B, w + "left", atts)
    ET.SubElement(wtc2B, w + "bottom", atts)
    ET.SubElement(wtc2B, w + "right", atts)
    ET.SubElement(wtc2Pr, w + "vAlign", attrib={w + "val": "center"})
    wtc2.append(wtc2Pr)
    
    for i in range(len(PS2)):
        index = i + 1
        if index == len(PS2):
            final = tbl6ParaEle(PS2[i])
            ET.SubElement(final, w + "bookmarkStart", attrib={w + "name": "_GoBack", w + "id": "0"})
            ET.SubElement(final, w + "bookmarkEnd", attrib={w + "id": "0"})
            wtc2.append(final)
        else:
            wtc2.append(tbl6ParaEle(PS2[i]))
    return tree
            
def wrElement(value='', wt=True):
    # make a "run" element with decent children.
    # return a "w:r" element.
    w = r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    
    wr = ET.Element(w + 'r')
    wrPr = ET.SubElement(wr, w + 'rPr')
    ET.SubElement(wrPr, w + 'rFonts', attrib={w + "hint": "eastAsia"})
    ET.SubElement(wrPr, w + 'sz', attrib={w + "val": "20"})
    ET.SubElement(wrPr, w + 'szCs', attrib={w + "val": "20"})
    if wt:
        ET.SubElement(wr, w + 't').text = value
    return wr

def tbl6ParaEle(value):
    # make the paragraph Element for table 6.
    # return a "w:p" element.
    w = r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    wp = ET.Element(w + "p", attrib={w + "rsidR": "00865CD2", w + "rsidRDefault": "00865CD2", w + "rsidP": "0046785D", w + "rsidRPr": "0046785D"})
    wpPr = ET.SubElement(wp, w + "pPr")
    wr = wrElement(value)
    wp.append(wr)
    ET.SubElement(wpPr, w + "pStyle", attrib={w + "val": "Web"})
    wnumPr = ET.Element(w + "numPr")
    ET.SubElement(wnumPr, w + "ilvl", attrib={w + "val": "0"})
    ET.SubElement(wnumPr, w + "numId", attrib={w + "val": "1"})
    wpPr.append(wnumPr)
    ET.SubElement(wpPr, w + "spacing", attrib={w + "after": "0", w + "afterAutospacing": "0"})
    wpPr.append(wrElement(wt=False).find(w + 'rPr'))
    return wp

def twoStuff(tbl, stuff, night): # write ["\d", "\d"] in the corresponding cells
    # 當日主管f1
    for i in range(2, 4):
        tbl[i-1][1].append(stuff[0])
    for i in range(4, 6):
        tbl[i-1][7].append(stuff[0])
    for i in range(6, 10):
        tbl[i-1][2].append(stuff[0])
    for i in range(10, 12):
        tbl[i-1][6].append(stuff[0])
    for i in range(12, 14):
        tbl[i-1][1].append(stuff[0])
    for i in range(14, 16):
        tbl[i-1][7].append(stuff[0])
    # f2
    for i in range(2, 6):
        tbl[i-1][2].append(stuff[1])
    for i in range(6, 10):
        tbl[i-1][7].append(stuff[1])
    for i in range(10, 16):
        tbl[i-1][2].append(stuff[1])
        
    if night == stuff[0]:
        notNight = stuff[1]
    else:
        notNight = stuff[0]
    for i in range(16, 26):
        tbl[i-1][1].append(night)
        tbl[i-1][2].append(notNight)
    return tbl

def threeStuff(tbl, stuff, night):
    if stuff[2] == night:
        notNight = stuff[1]
    else:
        night = stuff[1] 
        notNight = stuff[2]
    if "1" in stuff:
        # f1 = "1"
        for i in range(2, 4):
            tbl[i-1][1].append("1")
        for i in range(4, 10):
            tbl[i-1][7].append("1")
        for i in range(10, 12):
            tbl[i-1][6].append("1")
        for i in range(12, 14):
            tbl[i-1][1].append("1")
        for i in range(14, 26):
            tbl[i-1][7].append("1")
        # 不用值宿f2
        for i in range(2, 6):
            tbl[i-1][2].append(notNight)
        for i in range(6, 10):
            tbl[i-1][7].append(notNight)
        for i in range(10, 14):
            tbl[i-1][2].append(notNight)
        for i in range(14, 16):
            tbl[i-1][7].append(notNight)
        for i in range(16, 26):
            tbl[i-1][2].append(notNight)
        # 要值宿f3
        for i in range(2, 6):
            tbl[i-1][7].append(night)
        for i in range(6, 10):
            tbl[i-1][2].append(night)
        for i in range(10, 12):
            tbl[i-1][6].append(night)
        for i in range(12, 14):
            tbl[i-1][7].append(night)
        for i in range(14, 16):
            tbl[i-1][2].append(night)
        for i in range(16, 26):
            tbl[i-1][1].append(night)
    if "1" not in stuff:
        # 代理主管f1
        for i in range(2, 4):
            tbl[i-1][1].append(stuff[0])
        for i in range(4, 6):
            tbl[i-1][7].append(stuff[0])
        for i in range(6, 10):
            tbl[i-1][2].append(stuff[0])
        for i in range(10, 12):
            tbl[i-1][6].append(stuff[0])
        for i in range(12, 14):
            tbl[i-1][1].append(stuff[0])
        for i in range(14, 16):
            tbl[i-1][2].append(stuff[0])
        for i in range(16, 26):
            tbl[i-1][7].append(stuff[0])
        # 不用值宿f2
        for i in range(2, 6):
            tbl[i-1][2].append(notNight)
        for i in range(6, 10):
            tbl[i-1][7].append(notNight)
        for i in range(10, 12):
            tbl[i-1][6].append(notNight)
        for i in range(12, 16):
            tbl[i-1][7].append(notNight)
        for i in range(16, 26):
            tbl[i-1][2].append(notNight)
        # 要值宿f3
        for i in range(2, 10):
            tbl[i-1][7].append(night)
        for i in range(10, 14):
            tbl[i-1][2].append(night)
        for i in range(14, 16):
            tbl[i-1][7].append(night)
        for i in range(16, 26):
            tbl[i-1][1].append(night)
    return tbl

def twoSMS(tbl, SMS):
    # s1
    for i in range(2, 8):
        tbl[i-1][2].append(SMS[0])
    for i in range(8, 12):
        tbl[i-1][1].append(SMS[0])
    for i in range(12, 14):
        tbl[i-1][7].append(SMS[0])
    for i in range(14, 20):
        tbl[i-1][2].append(SMS[0])
    for i in range(20, 26):
        tbl[i-1][7].append(SMS[0])
    # s2
    for i in range(2, 4):
        tbl[i-1][7].append(SMS[1])
    for i in range(4, 8):
        tbl[i-1][1].append(SMS[1])
    for i in range(8, 14):
        tbl[i-1][2].append(SMS[1])
    for i in range(14, 16):
        tbl[i-1][1].append(SMS[1])
    for i in range(16, 20):
        tbl[i-1][7].append(SMS[1])
    for i in range(20, 26):
        tbl[i-1][2].append(SMS[1])
    return tbl
        
def threeSMS(tbl, SMS):
    # s1
    for i in range(2, 6):
        tbl[i-1][2].append(SMS[0])
    for i in range(6, 9):
        tbl[i-1][7].append(SMS[0])
    for i in range(9, 12):
        tbl[i-1][1].append(SMS[0])
    for i in range(12, 14):
        tbl[i-1][7].append(SMS[0])
    for i in range(14, 18):
        tbl[i-1][2].append(SMS[0])
    for i in range(18, 26):
        tbl[i-1][7].append(SMS[0])
    # s2
    for i in range(2, 4):
        tbl[i-1][7].append(SMS[1])
    for i in range(4, 6):
        tbl[i-1][1].append(SMS[1])
    for i in range(6, 10):
        tbl[i-1][2].append(SMS[1])
    for i in range(10, 12):
        tbl[i-1][6].append(SMS[1])
    for i in range(12, 14):
        tbl[i-1][7].append(SMS[1])
    for i in range(14, 16):
        tbl[i-1][1].append(SMS[1])
    for i in range(16, 18):
        tbl[i-1][7].append(SMS[1])
    for i in range(18, 22):
        tbl[i-1][2].append(SMS[1])
    for i in range(22, 26):
        tbl[i-1][7].append(SMS[1])
    # s3
    for i in range(2, 6):
        tbl[i-1][7].append(SMS[2])
    for i in range(6, 9):
        tbl[i-1][1].append(SMS[2])
    tbl[8][7].append(SMS[2])
    for i in range(10, 14):
        tbl[i-1][2].append(SMS[2])
    for i in range(14, 22):
        tbl[i-1][7].append(SMS[2])
    for i in range(22, 26):
        tbl[i-1][2].append(SMS[2])
    return tbl

def fourSMS(tbl, SMS):
    # s1
    for i in range(2, 5):
        tbl[i-1][2].append(SMS[0])
    tbl[4][7].append(SMS[0])
    for i in range(6, 8):
        tbl[i-1][1].append(SMS[0])
    for i in range(8, 10):
        tbl[i-1][7].append(SMS[0])
    for i in range(10, 12):
        tbl[i-1][6].append(SMS[0])
    for i in range(12, 14):
        tbl[i-1][7].append(SMS[0])
    for i in range(14, 16):
        tbl[i-1][1].append(SMS[0])
    for i in range(16, 20):
        tbl[i-1][7].append(SMS[0])
    for i in range(20, 23):
        tbl[i-1][2].append(SMS[0])
    for i in range(23, 26):
        tbl[i-1][7].append(SMS[0])
    # s2
    for i in range(2, 5):
        tbl[i-1][7].append(SMS[1])
    for i in range(5, 8):
        tbl[i-1][2].append(SMS[1])
    for i in range(8, 10):
        tbl[i-1][7].append(SMS[1])
    for i in range(10, 12):
        tbl[i-1][1].append(SMS[1])
    for i in range(12, 14):
        tbl[i-1][7].append(SMS[1])
    for i in range(14, 17):
        tbl[i-1][2].append(SMS[1])
    for i in range(17, 26):
        tbl[i-1][7].append(SMS[1])
    # s3
    for i in range(2, 4):
        tbl[i-1][7].append(SMS[2])
    for i in range(4, 6):
        tbl[i-1][1].append(SMS[2])
    for i in range(6, 8):
        tbl[i-1][7].append(SMS[2])
    for i in range(8, 11):
        tbl[i-1][2].append(SMS[2])
    tbl[10][6].append(SMS[2])
    for i in range(12, 17):
        tbl[i-1][7].append(SMS[2])
    for i in range(17, 20):
        tbl[i-1][2].append(SMS[2])
    for i in range(20, 26):
        tbl[i-1][7].append(SMS[2])
    # s4
    for i in range(2, 8):
        tbl[i-1][7].append(SMS[3])
    for i in range(8, 10):
        tbl[i-1][1].append(SMS[3])
    tbl[9][6].append(SMS[3])
    for i in range(11, 14):
        tbl[i-1][2].append(SMS[3])
    for i in range(14, 23):
        tbl[i-1][7].append(SMS[3])
    for i in range(23, 26):
        tbl[i-1][2].append(SMS[3])
    return tbl

def fourSMSEdu(tbl, SMS): 
    # s1
    for i in range(2, 6):
        tbl[i-1][2].append(SMS[0])
    for i in range(6, 8):
        tbl[i-1][1].append(SMS[0])
    for i in range(8, 10):
        tbl[i-1][7].append(SMS[0])
    for i in range(10, 12):
        tbl[i-1][6].append(SMS[0])
    for i in range(12, 14):
        tbl[i-1][7].append(SMS[0])
    for i in range(14, 16):
        tbl[i-1][1].append(SMS[0])
    for i in range(16, 26):
        tbl[i-1][7].append(SMS[0])
    # s2
    for i in range(2, 4):
        tbl[i-1][7].append(SMS[1])
    for i in range(4, 6):
        tbl[i-1][1].append(SMS[1])
    for i in range(6, 8):
        tbl[i-1][2].append(SMS[1])
    for i in range(8, 10):
        tbl[i-1][7].append(SMS[1])
    for i in range(10, 12):
        tbl[i-1][6].append(SMS[1])
    for i in range(12, 14):
        tbl[i-1][7].append(SMS[1])
    for i in range(14, 18):
        tbl[i-1][2].append(SMS[1])
    for i in range(18, 26):
        tbl[i-1][7].append(SMS[1])
    # s3
    tbl[6][7].append(SMS[2])
    for i in range(8, 10):
        tbl[i-1][2].append(SMS[2])
    for i in range(10, 12):
        tbl[i-1][1].append(SMS[2])
    for i in range(12, 18):
        tbl[i-1][7].append(SMS[2])
    for i in range(18, 22):
        tbl[i-1][2].append(SMS[2])
    for i in range(22, 26):
        tbl[i-1][7].append(SMS[2])
    # s4
    tbl[6][7].append(SMS[3])
    for i in range(8, 10):
        tbl[i-1][1].append(SMS[3])
    for i in range(10, 14):
        tbl[i-1][2].append(SMS[3])
    for i in range(14, 22):
        tbl[i-1][7].append(SMS[3])
    for i in range(22, 26):
        tbl[i-1][2].append(SMS[3])
    return tbl

def f2s1(tbl, stuff, night, SMS):
    # 當日主管f1
    for i in range(2, 8):
        tbl[i-1][1].append(stuff[0])
    for i in range(8, 16):
        tbl[i-1][2].append(stuff[0])
    # f2
    for i in range(2, 8):
        tbl[i-1][2].append(stuff[1])
    for i in range(8, 12):
        tbl[i-1][1].append(stuff[1])
    for i in range(12, 16):
        tbl[i-1][2].append(stuff[1])
    # 值宿night
    if night == stuff[0]:
        notNight = stuff[1]
    else:
        notNight = stuff[0]
    for i in range(16, 26):
        tbl[i-1][1].append(night)
    for i in range(16, 26):
        tbl[i-1][2].append(notNight)
    # SMS
    for i in range(2, 12):
        tbl[i-1][2].append(SMS[0])
    for i in range(12, 16):
        tbl[i-1][1].append(SMS[0])
    for i in range(16, 26):
        tbl[i-1][2].append(SMS[0])
    return tbl

def f3s1(tbl, stuff, SMS): # 找到的範本只有 f1="1" 的情況
    # f1 = '1'
    for i in range(2, 6):
        tbl[i-1][1].append(stuff[0])
    for i in range(6, 10):
        tbl[i-1][7].append(stuff[0])
    for i in range(10, 12):
        tbl[i-1][6].append(stuff[0])
    for i in range(12, 14):
        tbl[i-1][1].append(stuff[0])
    for i in range(14, 26):
        tbl[i-1][7].append(stuff[0])
    # 非值宿f2
    for i in range(2, 14):
        tbl[i-1][2].append(stuff[1])
    for i in range(14, 16):
        tbl[i-1][1].append(stuff[1])
    for i in range(16, 26):
        tbl[i-1][2].append(stuff[1])
    # 值宿f3
    for i in range(2, 6):
        tbl[i-1][7].append(stuff[2])
    for i in range(6, 9):
        tbl[i-1][2].append(stuff[2])
    for i in range(9, 12):
        tbl[i-1][1].append(stuff[2])
    for i in range(12, 16):
        tbl[i-1][2].append(stuff[2])
    for i in range(16, 26):
        tbl[i-1][1].append(stuff[2])
    # SMS
    for i in range(2, 6):
        tbl[i-1][2].append(SMS[0])
    for i in range(6, 9):
        tbl[i-1][1].append(SMS[0])
    for i in range(9, 12):
        tbl[i-1][2].append(SMS[0])
    for i in range(12, 14):
        tbl[i-1][7].append(SMS[0])
    for i in range(14, 26):
        tbl[i-1][2].append(SMS[0])
    return tbl

def f2s0(tbl, stuff, night):
    # '1'在
    if '1' in stuff:
        for i in range(2, 6):
            tbl[i-1][1].append(stuff[0])
        for i in range(6, 16):
            tbl[i-1][2].append(stuff[0])
        for i in range(16, 26):
            tbl[i-1][1].append(stuff[0])

        for i in range(2, 6):
            tbl[i-1][2].append(stuff[1])
        for i in range(6, 16):
            tbl[i-1][1].append(stuff[1])
        for i in range(16, 26):
            tbl[i-1][2].append(stuff[1])
    else:
        # '1'不在，代理主管值宿
        if night == stuff[0]:
            for i in range(2, 8):
                tbl[i-1][1].append(stuff[0])
            for i in range(8, 16):
                tbl[i-1][2].append(stuff[0])
            for i in range(16, 26):
                tbl[i-1][1].append(stuff[0])

            for i in range(2, 8):
                tbl[i-1][2].append(stuff[1])
            for i in range(8, 16):
                tbl[i-1][1].append(stuff[1])
            for i in range(16, 26):
                tbl[i-1][2].append(stuff[1])
        # '1'不在，代理主管非值宿
        if night == stuff[1]:
            for i in range(2, 12):
                tbl[i-1][1].append(stuff[0])
            for i in range(12, 26):
                tbl[i-1][2].append(stuff[0])

            for i in range(2, 12):
                tbl[i-1][2].append(stuff[1])
            for i in range(12, 26):
                tbl[i-1][1].append(stuff[1])
    return tbl

def f3s0(tbl, stuff, night):
    if night == stuff[2]:
        notNight = stuff[1]
    else:
        notNight = stuff[2]
    # 當日主管f1
    for i in range(2, 4):
        tbl[i-1][1].append(stuff[0])
    for i in range(4, 12):
        tbl[i-1][2].append(stuff[0])
    for i in range(12, 14):
        tbl[i-1][1].append(stuff[0])
    for i in range(14, 26):
        tbl[i-1][2].append(stuff[0])
    # 非值宿f2
    for i in range(2, 4):
        tbl[i-1][2].append(notNight)
    for i in range(4, 8):
        tbl[i-1][1].append(notNight)
    for i in range(8, 16):
        tbl[i-1][2].append(notNight)
    for i in range(16, 26):
        tbl[i-1][1].append(notNight)
    # 值宿f3
    for i in range(2, 4):
        tbl[i-1][2].append(night)
    for i in range(4, 8):
        tbl[i-1][2].append(night)
    for i in range(8, 12):
        tbl[i-1][1].append(night)
    for i in range(12, 14):
        tbl[i-1][2].append(night)
    for i in range(14, 16):
        tbl[i-1][1].append(night)
    for i in range(16, 26):
        tbl[i-1][2].append(night)
    return tbl

def allDayTrain(tbl, par): # 08-18 train, and 18-08(the next day) live outside
    for i in range(2, 12):
        for j in par:
            tbl[i-1][6].append(j)
    return tbl

def HDTrain2Stuff(tbl, stuff, par, night):
    # 當日主管參與半日訓練
    if par == stuff[0]:
        for i in range(2, 12):
            tbl[i-1][6].append(par)
            tbl[i-1][2].append(stuff[1])
        for i in range(12, 14):
            tbl[i-1][1].append(par)
            tbl[i-1][2].append(stuff[1])
        for i in range(14, 16):
            tbl[i-1][7].append(par)
            tbl[i-1][2].append(stuff[1])
    # 不是當日主管參與訓練
    else:
        for i in range(2, 12):
            tbl[i-1][6].append(par)
            tbl[i-1][2].append(stuff[0])
        for i in range(12, 14):
            tbl[i-1][1].append(stuff[0])
            tbl[i-1][2].append(par)
        for i in range(14, 16):
            tbl[i-1][7].append(stuff[0])
            tbl[i-1][2].append(par)
    # 值宿
    if night == stuff[0]:
        notNight = stuff[1]
    else:
        notNight = stuff[0]
    for i in range(16, 26):
        tbl[i-1][1].append(night)
        tbl[i-1][2].append(notNight)
    return tbl
        
def HDTrain3Stuff(tbl, stuff, par, night):
    st = list(stuff)
    st.remove(par)
    # 當日主管f1
    for i in range(2, 4):
        tbl[i-1][1].append(st[0])
    for i in range(4, 6):
        tbl[i-1][7].append(st[0])
    for i in range(6, 10):
        tbl[i-1][2].append(st[0])
    for i in range(10, 12):
        tbl[i-1][6].append(st[0])
    for i in range(12, 14):
        tbl[i-1][1].append(st[0])
    for i in range(14, 16):
        tbl[i-1][7].append(st[0])
    # 非值宿f2
    for i in range(2, 6):
        tbl[i-1][2].append(st[1])
    for i in range(6, 10):
        tbl[i-1][7].append(st[1])
    for i in range(10, 12):
        tbl[i-1][2].append(st[1])
    for i in range(12, 16):
        tbl[i-1][7].append(st[1])
    for i in range(16, 26):
        tbl[i-1][2].append(st[1])
    # 上午參與半日訓練人員
    for i in range(2, 12):
        tbl[i-1][6].append(par)
    for i in range(12, 16):
        tbl[i-1][2].append(par)
    # 值宿與非值宿
    if night == par:
        notNight = st[0]
    else:
        notNight = par
    for i in range(16, 26):
        tbl[i-1][1].append(night)
        tbl[i-1][7].append(notNight)
    return tbl

def numOrder(num):
    num.sort()
    val = ''
    for i in range(len(num)):
        if i == len(num) - 1:
            val = f'{val}{num[i]}'
        else:
            val = f'{val}{num[i]},'
    return val

def dataManager(tree, data, date):
    # table 2(日期)
    tree = tbl2Content(tree, data[0])
    # table 3(勤務時段班表)
    stuff = data[1].split(",")
    night = data[2]
    SMS = data[3].split(",")
    table3 = []
    for i in range(25):
        table3.append([])
        for j in range(9):
            table3[i].append([])
    if len(stuff) == 2:
        if '半天訓練' in data[6]:
            for i in data[6]:
                if re.match(r'(\d)番08-18時(火調|安檢)講習', i):
                    participant = re.match(r'(\d)番08-18時(火調|安檢)講習', i).group(1)
            HDTrain2Stuff(table3, stuff, participant, night)
        elif len(SMS) == 0:
            table3 = f2s0(table3, stuff, night)
        elif len(SMS) == 1:
            table3 = f2s1(table3, stuff, night, SMS)
        else:
            table3 = twoStuff(table3, stuff, night)
    elif len(stuff) == 3:
        if '半天訓練' in data[6]:
            for i in data[6]:
                if re.match(r'(\d)番08-18時(火調|安檢)講習', i):
                    participant = re.match(r'(\d)番08-18時(火調|安檢)講習', i).group(1)
            HDTrain3Stuff(table3, stuff, participant, night)
            data[6].remove('半天訓練')
        elif len(SMS) == 0:
            table3 = f3s0(table3, stuff)
        elif len(SMS) == 1:
            table3 = f3s1(table3, stuff, SMS)
        else:
            table3 = threeStuff(table3, stuff, night)
    if len(SMS) == 2:
        table3 = twoSMS(table3, SMS)
    if len(SMS) == 3:
        table3 = threeSMS(table3, SMS)
    # 全天訓練(外宿)
    for i in data[6]:
        if re.match(r'全天訓練(\d+)', i):
            participants = list(re.match(r'全天訓練(\d+)', i).group(1))
            table3 = allDayTrain(table3, participants)
            data[6].remove(re.match(r'全天訓練(\d+)', i).group())
    # 法紀教育
    if len(SMS) == 4 and '法紀教育' in data[6]:
        table3 = fourSMSEdu(table3, SMS)
        data[6].remove('法紀教育')
    elif len(SMS) == 4:
        table3 = fourSMS(table3, SMS)
    # 防溺與水查
    for i in data[6]:
        if re.match('防溺', i):
            tree = drwningProof(tree)
            table3[3][4] = list(table3[3][7].pop(0))
            table3[4][4] = list(table3[4][7].pop(0))
            data[6].remove('防溺')
        if re.match('水查', i):
            table3[3][5] = list(table3[3][7].pop(0))
            table3[4][5] = list(table3[4][7].pop(0))
            data[6].remove('水查')
    # 六、日沒有訓練
    if date[12] == '六' or date[12] == '日':
        table3[9][7] = table3[9][6]
        table3[10][7] = table3[10][6]
        table3[9][6] = ''
        table3[10][6] = ''
    # 只有兩警消上班，而其中一位要去半天訓練的時候
    if '半天訓練' in data[6] and len(stuff) == 2:
        data[6].remove('半天訓練')
        if len(SMS) > 2:
            table3[1][1].append(table3[1][7].pop(len(table3[1][7])-1))
            table3[2][1].append(table3[2][7].pop(len(table3[2][7])-1))
        elif len(SMS) == 2: # 唯有役男也只有兩人的時候備勤格才會空掉
            table3[1][1].append(table3[1][2].pop(len(table3[1][2])-1))
            table3[2][1].append(table3[2][2].pop(len(table3[2][2])-1))
            table3[1][2].append(table3[1][7].pop(len(table3[1][7])-1))
            table3[2][2].append(table3[2][7].pop(len(table3[2][7])-1))
        else:
            print('←這張勤務表不在範本裡，請額外關注它。')
    # 寫入xml
    for i in range(25):
        for j in range(9):
            try:
                table3[i][j] = numOrder(table3[i][j])
            except AttributeError:
                pass
            tree = tbl34Content(tree, i, j, table3[i][j], 3)
    # table 4(各種休假表格)
    tree = tbl34Content(tree, 0, 1, data[4][0], 4)
    tree = tbl34Content(tree, 0, 3, data[4][1], 4)
    tree = tbl34Content(tree, 0, 5, data[4][2], 4)
    tree = tbl34Content(tree, 1, 1, data[4][3], 4)
    tree = tbl34Content(tree, 1, 3, data[4][4], 4)
    tree = tbl34Content(tree, 1, 5, data[4][5], 4)
    # table 6(出動梯次、附記)
    tree = tbl6Content(tree, data[5], data[6])
    return tree

def archive(date, source, ps):
    # archive files under a directory into a (date).docx file under CWD
    if ps == '':
        with zipfile.ZipFile(f'{date}.docx', 'w', zipfile.ZIP_DEFLATED, True, 9) as docu:
            rlen = len(os.path.relpath(source)) + 1
            for root, dirs, files in os.walk(source):
                for file in files:
                    file = os.path.join(root, file)
                    docu.write(file, os.path.relpath(file)[rlen:])
    else:
        with zipfile.ZipFile(f'{date}－{ps}.docx', 'w', zipfile.ZIP_DEFLATED, True, 9) as docu:
            rlen = len(os.path.relpath(source)) + 1
            for root, dirs, files in os.walk(source):
                for file in files:
                    file = os.path.join(root, file)
                    docu.write(file, os.path.relpath(file)[rlen:])

def dateName(day, stuff, SMS, Cien=None): # Cien = 慈恩
    # input yyymmdd, output yyymmdd(_+_)□
    d = re.compile(r"(\d\d\d)(\d\d)(\d\d)")
    wkDay = ["一", "二", "三", "四", "五", "六", "日"]
    lenStuff = int((len(stuff)+1)/2)
    lenSMS = int((len(SMS)+1)/2)

    if Cien:
        number = f'({lenStuff+1}+{lenSMS})'
    else:
        number = f'({lenStuff}+{lenSMS})'

    mo = d.search(day)
    year = 1911 + int(mo.group(1))
    date = datetime.datetime(year, int(mo.group(2)), int(mo.group(3)))
    day = day + number + wkDay[date.weekday()]
    return day

def questAccepted(data):
    desk = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    SS = os.path.join(desk, f'{data[0][:3]}年{data[0][3:5]}月份勤務表')
    temp = os.path.join(SS, 'temp')
    word = os.path.join(temp, 'word')
    os.chdir(desk)
    try:
        os.mkdir(f'{data[0][:3]}年{data[0][3:5]}月份勤務表')
        os.chdir(SS)
    except FileExistsError:
        os.chdir(SS)
        
    with zipfile.ZipFile(r"C:\Users\Sleepylizard\Desktop\SSMaker\SSTemplate.docx") as docu: # path here is for temporary.###################################
        tree = ET.parse(docu.open('word/document.xml'))
        os.mkdir("temp")
        docu.extractall(temp)
        os.remove(os.path.join(word, "document.xml"))
        
    namePs = []
    if '支援慈恩' in data[6]:
        namePs.append(data[6].pop(data[6].index('支援慈恩')))
        date = dateName(data[0], data[1], data[3], Cien=True)
    else:
        date = dateName(data[0], data[1], data[3])
    if '法紀教育' in data[6]:
        namePs.append('法紀教育')
    if '常訓' in data[6]:
        namePs.append('常訓')
        data[6].remove('常訓')
    if 'T2訓' in data[6]:
        namePs.append('T2訓')
        data[6].remove('T2訓')
    if '水訓' in data[6]:
        namePs.append('水訓')
        data[6].remove('水訓')
    for i in data[6]:
        if re.match(r'(火調|安檢)講習', i):
            namePs.append(re.match(r'(火調|安檢)講習', i).group())
            data[6].remove(re.match(r'(火調|安檢)講習', i).group())
            data[6].append('半天訓練')
    namePs = '、'.join(namePs)

    misNote = [] # mistaken notation
    for i in data[6]:
        if re.match(r'(\d)番異常', i):
            errNum = re.match(r'(\d)番異常', i).group(1)
            data[6].remove(re.match(r'(\d)番異常', i).group())
            misNote.append(errNum)
    if len(misNote) == 0:
        pass
    elif len(misNote) == 1:
        print(f'←這張勤務表{errNum}番的輪休表註記不在範本裡，請額外給予關注。')
    elif len(misNote) > 1:
        misNoStr = misNote[0]
        for i in range(1, len(misNote)):
            misNoStr = f'{misNoStr}、{misNote[i]}'
        print(f'←這張勤務表{misNoStr}番的輪休表註記不在範本裡，請額外給予關注。')
            
    tree = dataManager(tree, data, date)
    os.chdir(word)
    xmlSetup(tree)
    os.chdir(SS)
    archive(date, temp, namePs)
    shutil.rmtree(temp)

def dataFetcher():
    # from .xlsx(輪休預定表)
    with zipfile.ZipFile(r'C:\Users\Sleepylizard\Desktop\輪休預定表108.01(4+4).xlsx') as xlsx: # path here is for temporily########################################
        si = [ele for eve, ele in ET.iterparse(xlsx.open('xl/sharedStrings.xml')) if ele.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si']
        labeledStrings = {} # {label_number:string}
        for i in range(len(si)):
            a = ''
            for text in si[i].itertext():
                a = a + text
            labeledStrings[i] = a
        elems = [elem for event, elem in ET.iterparse(xlsx.open('xl/worksheets/sheet1.xml'))]
        cell = {}
        for i in range(len(elems)):
            if elems[i].text:
                if 't' in elems[i+1].attrib:
                    if elems[i+1].attrib['t'] == 's': # need translate
                        cell[elems[i+1].attrib['r']] = labeledStrings[int(elems[i].text)]
                else: # need not to translate
                    cell[elems[i+1].attrib['r']] = str(elems[i].text)
        for i in range(8, 27, 2): # 預防備註是空的導致後續的KeyError
            try:
                cell[f'D{i}']
            except KeyError:
                cell[f'D{i}'] = ''
        del cell['D16']
                
    # from .docx(訓練預定表)
    with zipfile.ZipFile(r'C:\Users\Sleepylizard\Desktop\水璉分隊訓練預定表108.01.docx') as docx: # path here is for temporily###################################
        wTrs = [elem for event, elem in ET.iterparse(docx.open('word/document.xml')) if elem.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr']
        train = {} # {日期(1~31):['日期', '星期', '時段', '科目', '地點', '備註']}
        for i in range(2, len(wTrs)):
            tr = []
            for tc in wTrs[i].findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'):
                a = ''
                for text in tc.itertext():
                    a = f'{a}{text}'
                tr.append(a)
            train[i-1] = tr
    # make data list
    col = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH']
    data = {} # {day:[date, #1, #2, #3, #4, '', #5, #6, #7, #8, trainningSubject, PS, PS, PS, ...]}
    # make date string
    dayReg = re.compile(r'.*(\d\d\d).*(\d{1,2}).*')
    mo = dayReg.search(cell['A1'])
    year = str(mo.group(1))
    month = str(mo.group(2))
    if len(month) == 1:
        month = f'0{month}'
    date = year + month
    # arrange leave schedule and order notes
    for i in range(len(col)):
        data[i+1] = []
        day = i+1
        if i < 9:
            day = f'0{day}'
        data[i+1].append(date + str(day))
        for j in range(7, 24, 2):
            key = f'{col[i]}{j}'
            try :
                data[i+1].append(cell[key])
            except KeyError:
                data[i+1].append('')
    # arrange trainning schedule
    for i in range(31):
        if train[i+1][3] != '':
            data[i+1].append(f'16-18時{train[i+1][3]}')
        if train[i+1][5] != '':
            data[i+1].append(train[i+1][5])
    # leave notes of stuff(#1-D8, #2-D10, #3-D12, #4-D14)
    for i in range(8, 15, 2):
        pos = f'D{i}'
        for match in re.finditer(r'(\d|\d\d)日補?(\(補?(\d|\d\d)/(\d|\d\d)\s?\w+(-|－)?\w+\))', cell[pos]):
            data[int(match.group(1))].append(str(int(i/2-3)) + '番補休' + match.group(2))
        for match in re.finditer(r'(\d|\d\d)日事假?(\(.*\))', cell[pos]):
            data[int(match.group(1))].append(str(int(i/2-3)) + '番請事假' + match.group(2))
        for match in re.finditer(r'(\d|\d\d)日公假?(\(.*\))', cell[pos]):
            data[int(match.group(1))].append(str(int(i/2-3)) + '番請公假' + match.group(2))
        for match in re.finditer(r'(\d|\d\d)日病假?(\(.*\))', cell[pos]):
            data[int(match.group(1))].append(str(int(i/2-3)) + '番請病假' + match.group(2))
    # leave notes of SMS(#1-D8, #2-D10, #3-D12, #4-D14)
    for i in range(18, 25, 2):
        pos = f'D{i}'
        for match in re.finditer(r'(\d|\d\d)日補?(\(補?(\d|\d\d)/(\d|\d\d)\s?\w+(-|－)?\w+\))', cell[pos]):
            data[int(match.group(1))].append(str(int(i/2-4)) + '番補休' + match.group(2))
        for match in re.finditer(r'(\d|\d\d)日事假?(\(.*\))', cell[pos]):
            data[int(match.group(1))].append(str(int(i/2-4)) + '番請事假' + match.group(2))
        for match in re.finditer(r'(\d|\d\d)日公假?(\(.*\))', cell[pos]):
            data[int(match.group(1))].append(str(int(i/2-4)) + '番請公假' + match.group(2))
        for match in re.finditer(r'(\d|\d\d)日病假?(\(.*\))', cell[pos]):
            data[int(match.group(1))].append(str(int(i/2-4)) + '番請病假' + match.group(2))
    # final decoration
    if int(month) == 2:
        for i in range(29, 32):
            del data[i]
    if int(month) == 4  or int(month) == 6 or int(month) == 9 or int(month) == 11:
        del data[31]
    return data

def dataExplainer(data):
    a = [] # first data reassemble target: ['date', 'stuff', 'night', 'SMS', '輪休', '外宿', '補休', '休假', '公差假', '事病假', 'PS', ...]
    for day in data.keys():
        jar = {'stuff': [], 'night': [], 'SMS': [], 'tl': [], 'ol': [], 'cl': [], 'le': [], 'bl': [], 'pl': [], 'allDay':[]}
        olList = {'常訓': '', 'T2訓': 'EMT-2複訓', '水訓': '水域訓練'}
        # tl = 輪休, ol = 外宿, cl = 補休, le = 休假, bl = 公差假, pl = 事病假, allDay = 全天訓練參與人員
        SMSOrder = {} # {順位(int): 番號(str)}
        for i in range(1,5): # stuff
            if data[day][i] == '○':
                jar['tl'].append(str(i))
            elif re.match(r'春(\w)', data[day][i]):
                jar['tl'].append(str(i))
                ki = re.match(r'春(\w)', data[day][i]).group(1)
                data[day].append(f'{i}番春{ki}')
            elif data[day][i] == '特' or data[day][i] == '◎':
                data[day].append(f'{i}番特休')
                jar['tl'].append(str(i))
            elif data[day][i] == '補':
                jar['cl'].append(str(i))
            elif data[day][i] == '休':
                jar['le'].append(str(i))
                data[day].append(f'{i}番請休假')
            elif data[day][i] == '公':
                jar['bl'].append(str(i))
            elif data[day][i] == '事' or data[day][i] == '病':
                jar['pl'].append(str(i))
            elif data[day][i] == '喪':
                jar['pl'].append(str(i))
                data[day].append(f'{i}番請喪假')
            elif data[day][i] == '宿' or data[day][i] == 'v' or data[day][i] == 'V':
                jar['night'].append(str(i))
                jar['stuff'].append(str(i))
            elif data[day][i] == '慈':
                data[day].append(f'{i}番支援慈恩勤務')
                data[day].append('支援慈恩')
            elif re.match(r'(常訓|T2訓|T2|常|水訓|水)', data[day][i]):
                jar['ol'].append(str(i))
                jar['allDay'].append(str(i))
                outLeave = re.match(r'(常訓|T2訓|T2|常水訓|水)', data[day][i]).group()
                if not re.search(r'訓', outLeave):
                    outLeave = f'{outLeave}訓'
                trainName = olList[outLeave]
                data[day].append(f'{i}番08-18時{trainName}；18-08時外宿')
                if not outLeave in data[day][10:]:
                    data[day].append(outLeave)
            elif re.match(r'(安檢|火調)', data[day][i]):
                edu = re.match(r'(安檢|火調)', data[day][i]).group()
                jar['stuff'].append(str(i))
                if re.search(r'(宿|v|V)', data[day][i]):
                    jar['night'].append(str(i))
                data[day].append(f'{i}番08-18時{edu}講習')
                data[day].append(f'{edu}講習')
            elif data[day][i] == '':
                jar['stuff'].append(str(i))
            else:
                data[day].append(f'{i}番異常')
        for j in range(6, 10): # SMS
            try:
                if isinstance(int(data[day][j]), int):
                    SMSOrder[int(data[day][j])] = str(j-1)
            except ValueError:
                if data[day][j] == '○':
                    jar['tl'].append(str(j-1))
                elif re.match(r'(法紀|教育|法)(\d)', data[day][j]):
                    SMSOrder[int(re.match(r'(法紀|教育|法)(\d)', data[day][j]).group(2))] = str(j-1)
                    data[day].append(f'{j-1}番08-12時法紀教育')
                    if '法紀教育' not in data[day]:
                        data[day].append('法紀教育')
                elif data[day][j] == '喪':
                    jar['pl'].append(str(j-1))
                    data[day].append(f'{j-1}番請喪假')
                elif data[day][j] == '公':
                    jar['bl'].append(str(j-1))
                elif data[day][j] == '事' or data[day][j] == '病':
                    jar['pl'].append(str(j-1))
                elif data[day][j] == '': # 因應退役後的班表為空字串(這行好像沒意義?)
                    pass
                elif data[day][j] == '退役' or data[day][j] == '退':
                    data[day].append(f'{j-1}番退役')
                elif re.match(r'(常訓|T2訓|T2|常|水訓|水)', data[day][j]):
                    jar['ol'].append(str(j-1))
                    jar['allDay'].append(str(j-1))
                    outLeave = re.match(r'(常訓|T2訓|T2|常水訓|水)', data[day][j]).group()
                    if not re.search(r'訓', outLeave):
                        outLeave = f'{outLeave}訓'
                        trainName = olList[outLeave]
                        data[day].append(f'{i}番08-18時{trainName}；18-08時外宿')
                    if not outLeave in data[day][10:]:
                        data[day].append(outLeave)
                else:
                    data[day].append(f'{j-1}番異常')
        for k in range(len(SMSOrder.keys())):
            jar['SMS'].append(SMSOrder[k+1])
        if jar['allDay'] != []:
            retrain = ''.join(jar['allDay'])
            data[day].append(f'全天訓練{retrain}')
        del jar['allDay'] # 我懶惰改後面的東西了。刪掉它還原回沒問題的狀態比較輕鬆...
        a.append([])
        a[day-1] = [','.join(jar[l]) for l in jar.keys()]
        a[day-1].insert(0, data[day][0])  # now: ['date', 'stuff', 'night', 'SMS', '輪休', '外宿', '補休', '休假', '公差假', '事病假']
        # 出動梯次安排
        go = {'11車': [], '12車': [], '91車': [], '火值': []}
        go['91車'].append(a[day-1][1].split(',')[0]) # 當日主管
        go['11車'].append(a[day-1][1].split(',')[1]) # 非當日主管的最大番號
        if len(a[day-1][1].split(',')) > 2:
            go['12車'].append(a[day-1][1].split(',')[2]) # 第三位警消
        if 1 in SMSOrder.keys():
            go['91車'].append(SMSOrder[1])
        if 2 in SMSOrder.keys():
            go['火值'].append(SMSOrder[2])
        if 3 in SMSOrder.keys():
            go['11車'].append(SMSOrder[3])
        if 4 in SMSOrder.keys() and go['12車']:
            go['12車'].append(SMSOrder[4])
        elif 4 in SMSOrder.keys():
            go['火值'].append(SMSOrder[4])
        goString = ''
        for m in go.keys():
            if len(go[m]):
                goString = f'{goString}{m}：{go[m][0]}'
            elif m == '12車' or m == '火值':
                goString = f'{goString}{m}：'
            if len(go[m]) == 2:
                goString = f'{goString}、{go[m][1]}'
            goString = f'{goString}        '
        a[day-1].append(goString)
        # 車輛保養
        with open(r'C:\Users\Sleepylizard\Desktop\SSMaker\metadata.txt') as ref: # path...#############################################
            if a[day-1][0][5:] == '01':
                while True:
                    meta = ref.readline()
                    if meta:
                        yesterday = datetime.date(int(a[day-1][0][0:3])+1911, int(a[day-1][0][3:5]), int(a[day-1][0][5:])) + datetime.timedelta(-1)
                        if yesterday.month < 10:
                            mon = f'0{yesterday.month}'
                        else:
                            mon = str(yesterday.month)
                        yesterday = f'{int(yesterday.year)-1911}{mon}{yesterday.day}'
                        if meta[:7] == yesterday:
                            yStuff = meta.split('; ')[1]
                            break
            else:
                yStuff = a[day-2][1]
        carKeeper = []
        if '2' in yStuff or data[day][2] == '' or re.match(r'(宿|v)', data[day][2], re.IGNORECASE):
            carKeeper.append('2')
        if '3' in yStuff or data[day][3] == '' or re.match(r'(宿|v)', data[day][3], re.IGNORECASE):
            carKeeper.append('3')
        if '4' in yStuff or data[day][4] == '' or re.match(r'(宿|v)', data[day][4], re.IGNORECASE):
            carKeeper.append('4')  
        if '2' in carKeeper: # 91/72車
            ambu = '91/42車(2)'
        elif '2' not in carKeeper and '3' not in carKeeper:
            ambu = '91/42車(4)'
        elif '2' not in carKeeper:
            ambu = '91/42車(3)'
        if '3' in carKeeper: # 11車
            ft11 = '11車(3)'
        elif '3' not in carKeeper and '4' not in carKeeper:
            ft11 = '11車(2)'
        elif '3' not in carKeeper:
            ft11 = '11車(4)'
        if '4' in carKeeper: # 12車
            ft12 = '12車(4)'
        elif '4' not in carKeeper and '2' not in carKeeper:
            ft12 = '12車(3)'
        elif '4' not in carKeeper:
            ft12 = '12車(2)'
        keepers = f'{ft11}、{ft12}、{ambu}'
        a[day-1].append(keepers) # now: ['date', 'stuff', 'night', 'SMS', '輪休', '外宿', '補休', '休假', '公差假', '事病假', '出動梯次', '車輛保養']
        # 合併相同事由、番號不同的備註
        psInd = 10
        psReg = re.compile(r'(\d)(番.*)')
        for ps in data[day][10:]:
            psInd = psInd + 1
            mo = psReg.match(ps)
            for o in data[day][psInd:]:
                moo = psReg.match(o)
                if mo and moo and mo.group(2) == moo.group(2):
                    data[day].remove(ps)
                    data[day].remove(o)
                    ps = f'{mo.group(1)}、{moo.group(1)}{mo.group(2)}'
                    data[day].append(ps)
        for term in data[day][10:]:
            a[day-1].append(term) # ['date', 'stuff', 'night', 'SMS', '輪休', '外宿', '補休', '休假', '公差假', '事病假', '出動梯次', '車輛保養', 'PS', ...]
    output = [] # data reassemble target:　e.g. '1071126; 1,3; 1; 7,8,5,6; 2,4|||||; 11車：3、7        12車：        91車：1、5        火值：6、8; 車輛保養：11車(3) 、 12車(4) 、 91/72車(3)|16-18時搶救訓練－消防衣著裝|2番請休假'
    for p in a:
        tBox = ''
        lBox = ''
        for q in p[:4]:
            tBox = f'{tBox}{q}; '
        for r in p[4:9]:
            lBox = f'{lBox}{r}|'
        lBox = f'{lBox}{p[9]}'
        tBox = f'{tBox}{lBox}; {p[10]}; '
        lBox = ''
        for s in p[11:len(p)-1]:
            lBox = f'{lBox}{s}|'
        lBox = f'{lBox}{p[len(p)-1]}'
        tBox = f'{tBox}{lBox}'
        output.append(tBox)
    return output

if __name__ == '__main__':
    with open(r'C:\Users\Sleepylizard\Desktop\SSMaker\metadata.txt', 'ab') as meta: # path... ################################################
        for i in dataExplainer(dataFetcher()):
            print(f'製作{i[:3]}年{i[3:5]}月{i[5:7]}日的勤務表......', end='')
            j = codecs.encode(i + '\r\n', encoding='ANSI') # windows line-breaker is "\r\n"
            meta.write(j)
            i = i.replace('\n', '')
            i = i.split('; ')
            i[4] = i[4].split('|')
            i[6] = i[6].split('|')
            questAccepted(i)
            print('完成！')
    print('本月份勤務表全數製作完成！請務必再次確認。')
    os.system('pause >nul | echo 按下任何鍵以關閉程式...')
