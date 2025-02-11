import os
# 导入所有考试成绩文件名
def dir(file_dir):
    for root, dirs, files in os.walk(file_dir):
        return [dirs,files]

def file(file_dir):
    for root, dirs, files in os.walk(file_dir):
        pass
        # print('root:',root)  # 当前目录路径
        # print('dirs:',dirs)  # 当前路径下所有子目录
        # print('files:',files)  # 当前路径下所有非目录子文件'''
        return files
#将文件夹的名字转为对应编号
def rename_number(name):
    names=['教材回顾', '高考研究','单元检测','夯基保分练','提能增分练',"达标训练","课时对点练","例题"
        ,"实验","针对训练","微型专题练",'必修章']
    # names2=['作业']
    for i in range(len(names)):
        try:
            if  names[i] in name:
                if names[i]=='教材回顾':
                    number_1='J'
                elif names[i]=='高考研究':
                    number_1='G'
                elif names[i] == '单元检测':
                    number_1 = 'D'
                elif names[i] == '夯基保分练':
                    number_1 = 'H'
                elif names[i] == '提能增分练':
                    number_1 = 'T'
                elif names[i] == '达标训练':
                    number_1 = 'D'
                elif names[i] == '课时对点练':
                    number_1 = 'K'
                elif names[i] == '针对训练':
                    number_1 = 'Z'
                elif names[i] == '例题':
                    number_1 = 'L'
                elif names[i] == '微型专题练':
                    number_1 = 'W'
                elif names[i] == '实验':
                    number_1 = 'S'
                elif names[i] == '必修章':
                    number_1 = '50'
                number_2=name[len(names[i]):]
                number = str(number_1) + str(number_2)
        except:
            number=0
            pass

    # for i in range(len(names2)):
    #     try:
    #         if  names2[i] in name:
    #             number_1='Z'
    #             # print('name,number2:',name,number_1)
    #             number = str(number_1)
    #     except:
    #         pass
    # number=int(str(number_1)+str(number_2))

    return number

#将图片重命名并复制
def rename_picture(old_name,new_name,old_path,new_path):
    oldpath = old_path+'\\'+old_name
    newpath = new_path+'\\'+new_name+'.jpg'
    # print('oldpath:',oldpath,'newpath:',newpath)
    import shutil
    try:
        f = open(newpath)
        # print(newpath,'该文件存在,无法复制')
        f.close()
    except IOError:
        shutil.copy(oldpath, newpath)
        print(newpath,"完成复制")

#为题目图片重命名
def rename_title(number,title):
    title=title.strip('.jpg')
    try:
        title=int(title)
        if title < 10:
            title = '0' + str(title)
        # if title <100:
        #     title = '0'+str(title)
    except:pass
    title=number+str(title)
    # print(title)
    return title
if __name__=='__main__':
    path_all=[# r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第1单元','20190101',
              # r"F:\个性化题库20180423\title practice\word 题目\2019级创新设计\图片","2019",
              # r"F:\个性化题库20180423\title practice\图片 题目\5能量动量题目汇总 答案 图片","A2019",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级创新设计\答案","A2019",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级创新设计\答案","A2019",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节11","301901310111",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节11答案","A301901310111",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节21","301901310121",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节21答案","A301901310121",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节31","301901310131",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节31答案","A301901310131",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节41","301901310141",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节41答案","A301901310141",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节42","301901310142",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节42答案","A301901310142",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节51","301901310151",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节51答案","A301901310151",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节61","301901310161",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节61答案","A301901310161",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节71","301901310171",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节71答案","A301901310171",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节81","301901310181",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节81答案","A301901310181",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节91","301901310191",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节91答案","A301901310191",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节11","301901310211",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节21","301901310221",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节31","301901310231",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节41","301901310241",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节51","301901310251",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节61","301901310261",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节71","301901310271",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节81","301901310281",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节91","301901310291",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节101","3019013102101",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节42","301901310242",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节11答案","A301901310211",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节21答案","A301901310221",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节31答案","A301901310231",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节41答案","A301901310241",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节51答案","A301901310251",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节61答案","A301901310261",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节71答案","A301901310271",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节81答案","A301901310281",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节91答案","A301901310291",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节101答案","A3019013102101",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节42答案","A301901310242",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\节11","301901310111",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\微型专题1","3019013101W1",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\微型专题1答案","A3019013101W1",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\微型专题2","3019013101W2",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\微型专题2答案","A3019013101W2",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\微型专题3","3019013101W3",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\微型专题3答案","A3019013101W3",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\微型专题4","3019013101W4",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\微型专题4答案","A3019013101W4",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\微型专题5","3019013102W5",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\微型专题5答案","A3019013102W5",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\实验1","3019013102S1",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\实验1答案","A3019013102S1",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\实验2","3019013102S2",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第2章\实验2答案","A3019013102S2",

              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节11","301901310111",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节11答案","A301901310111",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节11","301901310111",
              # r"F:\个性化题库20180423\title practice\word 题目\2019级步步高选修3-1\第1章\节11答案","A301901310111",
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第1单元', '20190101',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第2单元', '20190102',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第3单元', '20190103',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第4单元', '20190104',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第5单元', '20190105',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第6单元', '20190106',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第7单元', '20190107',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第8单元', '20190108',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第9单元', '20190109',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第10单元', '20190110',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第11单元', '20190111',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第12单元', '20190112',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第13单元', '20190113',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第9单元 答案', 'A20190109',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第1单元 答案', 'A20190101',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第2单元 答案', 'A20190102',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第3单元 答案', 'A20190103',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第4单元 答案', 'A20190104',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第5单元 答案', 'A20190105',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第6单元 答案', 'A20190106',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第7单元 答案', 'A20190107',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第8单元 答案', 'A20190108',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第9单元 答案', 'A20190109',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第10单元 答案', 'A20190110',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第11单元 答案', 'A20190111',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第12单元 答案', 'A20190112',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第9单元 答案', 'A20190109',
              # r'F:\个性化题库20180423\title practice\word 题目\2019三维设计一轮复习\第13单元 答案', 'A20190113',
              # r'F:\考试\2016级\理综训练\第1次', '20190401',
              # r'F:\考试\2016级\理综训练\第2次', '20190402',
              # r'F:\考试\2016级\理综训练\第3次', '20190403',
              # r'F:\考试\2016级\理综训练\第4次', '20190404',
              # r'F:\考试\2016级\理综训练\第5次', '20190405',
              # r'F:\考试\2016级\理综训练\第6次', '20190406',
              # r'F:\考试\2016级\理综训练\答案\第1次', 'A20190401',
              # r'F:\考试\2016级\理综训练\答案\第2次', 'A20190402',
              # r'F:\考试\2016级\理综训练\答案\第3次', 'A20190403',
              # r'F:\考试\2016级\理综训练\答案\第4次', 'A20190404',
              # r'F:\考试\2016级\理综训练\答案\第5次', 'A20190405',
              # r'F:\考试\2016级\理综训练\答案\第6次', 'A20190406',
              #
              # r'F:\考试\2016级\大考\答案\高三上1考', '20190201',
              # r'F:\考试\2016级\大考\答案\高三上2考', '20190202',
              # r'F:\考试\2016级\大考\答案\高三上3考', '20190203',
              # r'F:\考试\2016级\大考\答案\高三上4考', '20190204',
              # r'F:\考试\2016级\大考\答案\高三上5考', '20190205',
              # r'F:\原题重做\试卷错题统计\历次考试（含图片和分值）\2016级高三下理综训练4衡水二调', '20194604',
              # r'F:\原题重做\试卷错题统计\历次考试（含图片和分值）\2016级高三下理综训练5衡水信息一', '20194605',
              # r'F:\原题重做\试卷错题统计\历次考试（含图片和分值）\2016级高三下理综训练6衡水信息二', '20194606',
              # r'F:\原题重做\试卷错题统计\历次考试（含图片和分值）\2016级高三下强化训练1衡水三', '20194611',
              # r'F:\原题重做\试卷错题统计\历次考试（含图片和分值）\2016级高三下强化训练2衡水四', '20194612',
              # r'F:\个性化题库20180423\title practice\图片 题目\牛顿运动定律 各资料汇总 图片','201625102',
              # r'F:\个性化题库20180423\title practice\图片 题目\1匀变速直线运动','20190301',
              # r'F:\个性化题库20180423\title practice\图片 题目\1匀变速直线运动 答案','A20190301',
              # r'F:\个性化题库20180423\title practice\图片 题目\2牛顿运动定律 各资料汇总 图片','20190302',
              # r'F:\个性化题库20180423\title practice\图片 题目\4万有引力','20190304',
              # r'F:\个性化题库20180423\title practice\图片 题目\3曲线运动','20190303',
              # r'F:\个性化题库20180423\title practice\图片 题目\3曲线运动 答案','A20190303',
              # r'F:\个性化题库20180423\title practice\图片 题目\5能量动量题目汇总 答案 图片','A20190305',
              # r'F:\个性化题库20180423\title practice\图片 题目\6力学实验','20190306',
              # r'F:\个性化题库20180423\title practice\图片 题目\9磁场','20190309',
              # r'F:\个性化题库20180423\title practice\图片 题目\7电场','20190307',
              # r'F:\个性化题库20180423\title practice\图片 题目\6力学实验 答案','A20190306',
              # r'F:\个性化题库20180423\title practice\图片 题目\9磁场 答案','A20190309',
              # r'F:\个性化题库20180423\title practice\图片 题目\4万有引力 答案','A20190304',
              # r'F:\个性化题库20180423\title practice\图片 题目\11交变电流','20190311',
              # r'F:\个性化题库20180423\title practice\图片 题目\11交变电流 答案','A20190311',
              # r'F:\个性化题库20180423\title practice\图片 题目\12近代物理 答案','A20190312',
              # r'F:\个性化题库20180423\title practice\图片 题目\5能量动量题目汇总 答案 图片','A20190305',
              # r'F:\个性化题库20180423\title practice\图片 题目\7电场 答案','A20190307',
              # r'F:\个性化题库20180423\title practice\图片 题目\2020级状元桥必修1\第二章','2020102',
              # r'F:\个性化题库20180423\title practice\图片 题目\2020级状元桥必修1\第三章','2020103',
              r'F:\个性化题库20180423\title practice\图片 题目\2020年秋状元桥必修2 - 副本\第五章','2020105',
              r'F:\个性化题库20180423\title practice\图片 题目\2020年秋状元桥必修2 - 副本\第五章答案','A2020105',
              # r'F:\个性化题库20180423\title practice\图片 题目\2020级状元桥必修1\第二章答案','A2020102',
              # r'F:\个性化题库20180423\title practice\图片 题目\12近代物理','20190312',
              # r'F:\个性化题库20180423\title practice\图片 题目\能量动量题目汇总 图片','20190305',
              # r'F:\个性化题库20180423\title practice\图片 题目\牛顿运动定律 各资料汇总选择题 图片','20190402',
              # r'F:\个性化题库20180423\title practice\图片 题目\恒定电流题目汇总 图片','20190308',
              # r'F:\个性化题库20180423\title practice\图片 题目\13热学题目汇总 图片','20190313',
              # r'F:\个性化题库20180423\title practice\图片 题目\13热学题目汇总 答案 图片','A20190313',
              # r'F:\个性化题库20180423\title practice\图片 题目\2牛顿运动定律 各资料汇总 答案 图片','A20190302',
              # r'F:\个性化题库20180423\title practice\图片 题目\恒定电流题目汇总 答案 图片','A20190308',
              # r'F:\个性化题库20180423\title practice\图片 题目\恒定电流题目汇总 答案 图片','A20190308',
              # r'F:\个性化题库20180423\title practice\个人文件\程明宇\picture',"",
              # r'F:\个性化题库20180423\title practice\图片 题目\牛顿运动定律 2019全品一轮复习 图片','20190202'
              ]
    # main_path=r'F:\原题重做\题库\2019三维设计一轮复习\第2单元' #201901
    # main_path=r'F:\个性化题库20180423\title practice\图片 题目\牛顿运动定律 各资料汇总 图片'#201625102
    # main_path=r'F:\个性化题库20180423\title practice\图片 题目\牛顿运动定律 2019全品一轮复习 图片'#201902
    # new_path = main_path  # 测试用
    for i in range(int(len(path_all)/2)):
        main_path=path_all[2*i]
        new_path = r'F:\个性化题库20180423\title practice\汇总'
        dirs = dir(main_path)
        number_0=path_all[2*i+1]
        # print('main_path,number_0:',main_path,number_0)
        # number_0 = '20190102'
        # number_0='20190101'
        # print(dirs)
        # print(i)
        title_all = []
        for k in dirs[1]:  # 第一层级下的题目处理
            title = rename_title(number_0, k)
            old_path = main_path
            # print('old_math', old_path)
            rename_picture(k, title, old_path, new_path)
            title_all.append(title)
        for i in dirs[0]:  # 第一层级，单元
            # print('i:', i, type(i))
            # number_1=0
            number_1 = rename_number(i)
            number = number_0 + number_1
            # number = 0
            # for k in dirs[1]:#第一层级下的题目处理
            #     title=rename_title(k)
            #     title_all.append(title)
            dirs2 = dir(r'%s\%s' % (main_path, i))
            for k in dirs2[1]:  # 第二层级下的题目处理
                title = rename_title(number, k)
                old_path = main_path + '\\' + i
                # print('old_math',old_path)
                rename_picture(k, title, old_path, new_path)
                title_all.append(title)
        print('title_all:', len(title_all), title_all)
        # for i in title_all:
            # print(i)
