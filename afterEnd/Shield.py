import os
import pathlib


def main(MprInput,MprOutput,IdList):
    List = []
    # mpr_list = os.listdir(MprInput)
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

    return True

#
# os.remove('37r.txt')
# os.rename('37r_swap.txt', '37r.txt')
# print('done...')
# done...

