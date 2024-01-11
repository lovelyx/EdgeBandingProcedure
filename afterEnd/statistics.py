import os
from xml.dom import minidom
import pandas as pd


def main(XmlOutput):

    FolderPath = XmlOutput
    xml_list = os.listdir(FolderPath)
    PanelList = []

    for j in xml_list:  # 遍历所有xml文件
        Sum = 0
        num = 0
        CNum = 0
        ANum = 0
        WuJgNum = 0
        fullPath = FolderPath + "/" + j  # 完整路径
        # 打开文件
        XmlDoc = minidom.parse(fullPath)
        # 找到Machining节点（标签）
        PanelNodes = XmlDoc.getElementsByTagName("Panel")
        print(j)
        for panel in PanelNodes:
            Sum += 1
            Machining = panel.getElementsByTagName("Machining")
            Machines2 = panel.getElementsByTagName("Machines")
            Length = panel.getAttribute('Length')
            width = panel.getAttribute('Width')
            Info7 = str(panel.getAttribute('Info7'))
            Info6 = str(panel.getAttribute('Info6'))
            if '5' in Info7 or '6' in Info7:
                CNum += 1
            if 'A' in Info6:
                ANum += 1
            if float(Length) < 250 or float(width) < 250:
                continue
            for mach in Machining:
                Type = mach.getAttribute('Type')

                if Type == '4':
                    num += 1
                    break
            if Machines2:
                pass
            else:
                WuJgNum += 1
        print(num)
        print(Sum)
        print(f'WuJgNum:{WuJgNum}')
        panel = [j, Sum, ANum, "{:.2f}%".format((ANum / Sum) * 100), num, "{:.2f}%".format((num / Sum) * 100), CNum]
        if num == 0:
            panel.append(0)
        else:
            panel.append("{:.2f}%".format((CNum / num)*100))
        PanelList.append(panel)
        # PanelList.append(WuJgNum)
    df = pd.DataFrame(PanelList)
    df.columns = ['批次号', '批次工件数量', 'A板', 'A板占比', '槽数量', '槽占比', '封边机开槽数量', '占比']
    path = FolderPath + '/' + "汇总" + '.csv'
    # test.save(path)
    df.to_csv(path, index=False, header=True, encoding='gb18030')

    return 1
