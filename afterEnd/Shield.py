import os
import pathlib
import logging
import traceback
import pymsgbox as mb


def main(MprInput,MprOutput,IdList,IdDictNo):
    try:
        List = []
        # mpr_list = os.listdir(MprInput)
        # print(IdDictNo)
        for j in IdList:
            for p in j:
                List.append(p)

        for root, dirs, files in os.walk(MprInput):
            # for file in files:
            #     path = os.path.join(root, file)
            #     print(path)
            for i in files:
                Mprfile = i.split(".")[0]
                if Mprfile in List and (i.split(".")[1] == 'MPR' or i.split(".")[1] == 'mpr'):
                    path = os.path.join(root, i)
                    # h1=0
                    # w=0
                    # L=0
                    # num = 0
                    # statX=0
                    # statY=0
                    # EndX=0
                    # EndY=0
                    with open(path) as fr:
                        # if os.path.exists(MprOutput+"/"+Mprfile.split('\\')[):
                        #     filepatch = MprOutput+"/"+Mprfile.split('\\',1)[-1]
                        # else:
                        #     os.mkdir(MprOutput+"/"+Mprfile.split('\\')[-2])
                        #     filepatch = MprOutput+"/"+Mprfile.split('\\',1)[-1]
                        # print(filepatch)
                        # print(path)
                        filepatch = MprOutput + "/" + path.split('\\')[-2] +"/" +path.split('\\')[-1]
                        filepatch2 = MprOutput + "/" + path.split('\\')[-2]
                        pathlib.Path(filepatch2).mkdir(parents=True, exist_ok=True)
                        # print(path.split('/')[-1])
                        with open(filepatch, 'w') as fw:
                            # 循环读取文件内容，逐行修改
                            # print(MprInput+"/"+Mprfile)
                            # print(oupFile+"/"+Mprfile)
                            for line in fr:
                                # if line.startswith('L'):
                                #     L=line.split('"')[1]
                                # if line.startswith('W'):
                                #     w=line.split('"')[1]
                                # if line.startswith(']'):
                                #     num = line.split(']')[1]
                                #     h1=2
                                # if line == '$E0' and h1 == 2:
                                #     h1=3
                                # if line.startswith('.X') and h1 ==3:
                                #     statX = line.split('=')[1]
                                # if line.startswith('.Y') and h1 ==3:
                                #     statY = line.split('=')[1]
                                #     if int(w)-int(statY) == 13.5:
                                #         h1=4
                                # if line == '$E1' and h1==4:
                                #     h1=5
                                # if line.startswith('.X') and h1 ==5:
                                #     EndX=line.split('=')[1]
                                #     if abs(int(EndX)-int(statX)) == L:
                                #         h1=6
                                # if line.startswith('.Y') and h1 == 6:
                                #     EndY = line.split('=')[1]
                                #     if int(statY) == int(EndY):
                                #         h1=7
                                # print(line)
                                if line.startswith("<105"):
                                    # h1=8
                                    # if h1 ==8 and line == "TNO='13'":
                                    #     h1=9
                                    # if h1 == 9 and line == "ZA='12'":
                                    #     h1 = 10
                                    # if h1==10 and line == "KAT='Router'":
                                    #     h1 =11
                                    # if h1==11:
                                    line = line + 'EN="0"\n'
                                    # h1 = 0
                                fw.write(line)
                elif (i.split(".")[1] == 'MPR' or i.split(".")[1] == 'mpr'):
                    path = os.path.join(root, i)
                    with open(path) as fr2:
                        filepatch = MprOutput + "/" + path.split('\\')[-2] +"/"+path.split('\\')[-1]
                        filepatch2 = MprOutput + "/" + path.split('\\')[-2]
                        pathlib.Path(filepatch2).mkdir(parents=True, exist_ok=True)
                        with open(filepatch, 'w') as fw2:
                            for line in fr2:
                                fw2.write(line)
            #
            # if IdDictNo!='' and root !='' and len(files) != 0:
            #     # print(root)
            #     # print(files)
            #     # print(dirs)
            #     path = root.split('\\')[-1]
                # print(path)
        lines = [
    '[H\n',
    'VERSION="4.0 Alpha"\n',
    'HP="1"\n',
    'IN="0"\n',
    'GX="0"\n',
    'BFS="1"\n',
    'GY="1"\n',
    'GXY="0"\n',
    'UP="0"\n',
    'FM="1"\n',
    'FW="800"\n',
    'HS="0"\n',
    'OP="2"\n',
    'MAT="WEEKE"\n',
    'INCH="0"\n',
    'View = "NOMIRROR"\n',
    'ANZ="1"\n',
    'BES="0"\n',
    'ENT="0"\n',
    '_BSX=1795.00\n',
    '_BSY=400.00\n',
    '_BSZ=18.00\n',
    '_FNX=0\n',
    '_FNY=0\n',
    '_RNX=0\n',
    '_RNY=0\n',
    '_RNZ=0\n',
    '_RX=1795.00\n',
    '_RY=400.00\n',
    '\n',
    '[001\n',
    'L="1795.00"\n',
    'KM="length"\n',
    'W="400.00"\n',
    'KM="width"\n',
    'T="18.00"\n',
    'KM="thickness"\n',
    '\n',
    '<100 \WerkStck\\\n',
    'LA="L"\n',
    'BR="W"\n',
    'DI="T"\n',
    'AX="0"\n',
    'AY="0"\n',
    'FNX="0"\n',
    'FNY="0"\n',
    '<101 \Kommentar\\\n',
    'KM="左侧板 1795.00*400.00*18.00"\n',
    '\n',
    '\n',
    '\n'
    '!'
    ]
        for i,j in IdDictNo.items():
            filepatch2=MprOutput + "/"+j.split('*')[0]
            # pathlib.Path(filepatch2).mkdir(parents=True, exist_ok=True)
            with open(f"{filepatch2}\\1{i}5.MPR", 'w', encoding='UTF-8') as f:
                Length = (j.split(' ')[1]).split('*')[0]
                width = (j.split(' ')[1]).split('*')[1]
                think = (j.split(' ')[1]).split('*')[2]
                lines[19] = f"_BSX={Length}\n"
                lines[20] = f"_BSY={width}\n"
                lines[21] = f"_BSZ={think}\n"
                lines[27] = f"_RX={Length}\n"
                lines[28] = f"_RY={width}\n"
                lines[31] = f'L="{Length}"\n'
                lines[33] = f'W="{width}"\n'
                lines[35] = f'T="{think}"\n'
                lines[47] = f'KM="{j}"\n'
                # print(lines)
                f.writelines(lines)
    # print(os.path.join(root, i))


        return True

    except:
        logging.basicConfig(filename='log.txt', level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        # 方案一，自己定义一个文件，自己把错误堆栈信息写入文件。
        errorFile = open('log.txt', 'a')
        errorFile.write(traceback.format_exc())
        errorFile.close()
        # yield f"{j}文件处理失败，请联系管理员"
        # tkinter.messagebox.showerror(title='错误',message=f"{j.split('.')[0]}文件处理失败，请联系管理员")
        mb.alert(f"mpr文件处理失败，请联系管理员", '报错')
        # win32api.MessageBox(0,f"{j.split('.')[0]}文件处理失败，请联系管理员","错误",win32con.MB_OK)




#
# os.remove('37r.txt')
# os.rename('37r_swap.txt', '37r.txt')
# print('done...')
# done...

