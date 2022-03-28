print("""
@Author:w01f
@repo:https://github.com/W01fh4cker/nucleic-acid-information-query
@version 1.0
@2022/3/27
___          ______          ___        ________         ____          _______
\  \        /  __  \        /  /      /   ____  \       |   |         /  ____/
 \  \      /  /  \  \      /  /       |  |   |  |       |   |     __/   /___
  \  \    /  /    \  \    /  /        |  |   |  |       |   |    /___   ___/
   \  \  /  /      \  \  /  /         |  |   |  |       |   |       |  |
    \  \/  /        \  \/  /          |  |___|  |       |   |       |  |
     \____/          \____/           \ ________/       |___|       |__|
""")
import requests
import xlrd
import xlwt
import json
import time

def status1():
    print("[*]" + name + "该人员信息有误")
    status1 = "该人员信息有误"
    output_worksheet.write(i, 3, status1)

def status2():
    print("[*]" + name + "该人员还没进行核酸检测！")
    status2 = "该人员还没进行核酸检测！"
    output_worksheet.write(i, 3, status2)

def status3():
    print("[*]" + name + "核酸检测报告还未出来")
    status3 = "核酸检测报告还未出来"
    output_worksheet.write(i, 3, status3)

def find_time(timechuo,output_worksheet,i):
    if timechuo is None:
        status3()
    else:
        global report_time_in
        report_time_in = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(int(timechuo)))
    output_worksheet.write(i, 1, report_time_in)

def find_result(list1,name,report_time_in):
    report_result = list1[0]["resultdata"]
    global status5
    status5 = "无"
    if (report_result == "2"):
        print("[*]" + name + "为阴性，核酸检测报告出具时间为：" + str(report_time_in))
    else:
        print("[*]" + name + "为阳性！！！！！！！！！，核酸检测报告出具时间为：" + str(report_time_in))
    output_worksheet.write(i, 2, report_result)
    output_worksheet.write(i, 3, status5)

def is_multi_relation(res,name):
    for j in range(len(res["list"])-1,-1,-1):
        # 公司版
        if(res["list"][j]["username"] == name):
            global newlist
            newlist = []
            newlist.append(res["list"][j])
        else:
            pass
    return newlist

def import_form_and_query():
    filepath = input("请输入您想要操作的excel表格的绝对路径(完整的)：")
    workbook = xlrd.open_workbook(filepath)  # 获取表格
    sheet_name = input("请输入您想要操作的表格文件的sheet名(左下角可以看到，注意大小写)：")
    work_sheet = workbook.sheet_by_name(sheet_name)  # 获取sheet名
    nrows = work_sheet.nrows - 1  # 统计当前sheet一共多少人
    print('共' + str(nrows) + '人。')  # 输出人数
    global i
    i = 1
    for i in range(1,nrows):
        global name
        name = work_sheet.cell_value(i, 0)  # 获取第二列（姓名）的内容
        output_worksheet.write(i, 0, name)
        sfz = '%s' % work_sheet.cell_value(i, 3)  # 获取第四列（身份证）的内容
        dict = {"username": name, "userid": str(sfz)} # 字典保存数据
        # 构造url
        url = 'https://cydj.weiynet.cn/api/v3/getUserResult?info=' + str(json.dumps(dict,ensure_ascii=False)) + '&queryType=1&usercode=undefined&phone=undefined&source=h5&version=v3'
        # 构造headers
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36"
        }
        # 构造请求
        resp = requests.post(url=url, headers=headers)
        # 接收json数据
        global res
        res = json.loads((resp.content).decode('utf-8'))
        # 判断信息是否有误
        if (res["err"] == -3):
            status1()
        elif (len(res["list"]) == -2):
            print("[*]请手动核查！！！！！！！！！")
        else:
            # 没做核酸
            if(len(res["list"]) == 0):
                status2()
            elif(len(res["list"]) == 1):
                timestamp1 = res["list"][0]["testtime"]
                find_time(timestamp1,output_worksheet,i)
                global res_1
                res_1 = res["list"]
                find_result(res_1,name,report_time_in)
            else:
                new_list = is_multi_relation(res,name)
                timestamp1 = new_list[0]["testtime"]
                if timestamp1 is None:
                    status3()
                else:
                    find_time(timestamp1,output_worksheet,i)
                    find_result(new_list,name,report_time_in)

def make_form():
    global output_workbook,output_worksheet
    output_workbook = xlwt.Workbook(encoding='utf-8')
    output_worksheet = output_workbook.add_sheet('核酸结果导出')
    # 创建颜色
    pattern = xlwt.Pattern()  # 创建模式对象
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5  # 设置模式颜色为黄色
    style = xlwt.XFStyle()  # 创建样式对象
    style.pattern = pattern  # 将模式加入到样式对象
    # 设置单元格的宽度
    output_worksheet.col(1).width = 600 * 20
    output_worksheet.col(2).width = 400 * 20
    output_worksheet.col(3).width = 600 * 20
    # 写第一行的标题
    output_worksheet.write(0, 0, '姓名', style)
    output_worksheet.write(0, 1, '最近一次核酸报告时间', style)
    output_worksheet.write(0, 2, '最近一次核酸检测结果(1表示阳性，2表示阴性)', style)
    output_worksheet.write(0, 3, '备注', style)

def main():
    make_form()
    import_form_and_query()
    localtime = time.localtime(time.time())
    make_time = time.strftime("%Y%m%d%H%M%S", time.localtime())
    output_workbook.save(make_time + '.xls')
    return again()
def again():
    while True:
        again = input("[*]本次已经检测完毕，如果想要继续请输入y，退出请按n。")
        if again not in {"y", "n"}:
            print("[*]请输入y或n而不是其他字符，谢谢！")
        elif again == "n":
            return "[*]感谢您的使用！"
        elif again == "y":
            return main()
if __name__ == '__main__':
    main()