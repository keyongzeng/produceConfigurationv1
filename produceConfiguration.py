import xlrd
import os
import time
import re
import zipfile
import sys

#传入excel表对象和变量列表
def replace_info(excel_name,bianliang_list):
    all_data = ""
    for j in range(1, 1000):
        try:
            command = excel_name.cell(j, 1).value

            if command == "EOF":
                break
            if re.search(r"&#\d+", command):
                try:
                    k = re.findall(r"&#\d+", command)
                    for row in k:
                        index = int(row.strip("&#"))
                        text = bianliang_list[index-1]
                        if type(text) == float:
                            text = str(int(text))
                        command = command.replace(row, text)
                except IndexError:
                    print("无&#%s参数不够,请补充所有参数后运行" % index)
                    sys.exit()
            all_data += command + "\n"
        except IndexError:
            return all_data
    return all_data

timefilename = time.strftime('%Y_%m_%d_%H%M')
if not os.path.exists("configration_storehouse"):
    os.mkdir("configration_storehouse")
while True:
    project = input("请输入本次工程名：").strip()
    if len(project) == 0:
        print("请输出工程号")
        continue
    dir_name_zip = "configration_storehouse/%s_%s_zip" % (project,timefilename)
    dir_name_txt = "configration_storehouse/%s_%s_txt" % (project,timefilename)
    if os.path.exists(dir_name_zip):
        print("操作太频繁，请一分钟后再操作。。。。")
    else:
        os.mkdir(dir_name_zip)
        os.mkdir(dir_name_txt)
        break

data = xlrd.open_workbook(r"信息表.xlsx")
info_table = data.sheet_by_name(r"参数库")
zip_table = data.sheet_by_name(r"生成压缩文件")
txt_table = data.sheet_by_name(r"生成txt文件")

for i in range(1,10000):
    try:
        flag = info_table.cell(i,0).value
        if flag == "N":
            continue
        zip_name = info_table.cell(i,1).value
        txt_name = info_table.cell(i, 2).value
        bianliang1 = info_table.cell(i, 3).value
        bianliang2 = info_table.cell(i, 4).value
        bianliang3 = info_table.cell(i, 5).value
        bianliang4 = info_table.cell(i, 6).value
        bianliang5 = info_table.cell(i, 7).value
        bianliang6 = info_table.cell(i, 8).value
        bianliang7 = info_table.cell(i, 9).value
        bianliang8 = info_table.cell(i, 10).value
        bianliang9 = info_table.cell(i, 11).value
        bianliang10 = info_table.cell(i, 12).value
        bianliang11 = info_table.cell(i, 13).value
        bianliang12 = info_table.cell(i, 13).value
        bianliang_list = [bianliang1,bianliang2,bianliang3,bianliang4,bianliang5,bianliang6,bianliang7,bianliang8,bianliang9,bianliang10,bianliang11,bianliang12]
        #print(bianliang_list)

        #压缩文件生成
        if len(zip_name.strip()) != 0:
            all_data = replace_info(zip_table,bianliang_list)
            if len(all_data.strip()) != 0:
                with open('temp/vrpcfg.cfg',"w") as f:
                    f.write(all_data)
                zip_filename = "%s/%s.zip" % (dir_name_zip,zip_name)
                z = zipfile.ZipFile(zip_filename,"w")
                z.write("temp/vrpcfg.cfg","vrpcfg.cfg")
                z.close()
            else:
                print("压缩表没有需要生成的命令")

        #文本文件生成
        if len(txt_name.strip()) != 0:
            to_data = replace_info(txt_table, bianliang_list)
            if len(to_data.strip()) != 0:

                txt_filename = "%s/%s.txt" % (dir_name_txt,txt_name)
                with open(txt_filename,"a") as file:
                    file.write(to_data)
            else:
                print("文本表没有需要生成的命令")

    except IndexError:
        break


