import threading
import requests
import pymysql
import xlwt


def main():
    print("start......")

    # token不定时获取
    # token = '1560976019693678_6253319_1690125190_2064a5c76c7381aaf32e59519060acbe'
    token = input("请输入token:")
    # 每echo轮控制台输出一次进度
    echo = 100
    # 每个excel存放女嘉宾信息数量
    rows = 500
    # excel表序号
    num = 0
    # 女嘉宾信息列表
    girl_list = []
    # 起始用户id
    fuid_begin = 0
    # 结束用户id
    fuid_end = 6543620  # 一周CP目前用户数量达到6543620
    # 抛异常的用户ID列表
    e_fuid = []

    thread_list = []
    for i in range(99):
        fuid_end = fuid_begin + 66100
        thread_list.append(threading.Thread(target=get_girls_to_db, args=(token, rows, fuid_begin, fuid_end, e_fuid)))
        fuid_begin = fuid_end
    for thread in thread_list:
        thread.start()
    for thread in thread_list:
        thread.join()

    print("efuid:", e_fuid)
    file = open('efuid.txt', 'w', encoding='utf-8')
    for i in range(len(e_fuid)):
        file.write(str(e_fuid[i]) + '\n')
    file.close()
    print("落表完毕")


# 爬取女嘉宾信息并记录数据库
def get_girls_to_db(token, rows, fuid_begin, fuid_end, e_fuid):
    girl_list = []
    for fuid in range(fuid_begin, fuid_end):
        try:
            # 获取用户信息
            resp = get_user_info(token, fuid)

            if resp["data"] and resp["data"]["sex_des"] == "女":  # resp["data"]["address"] == "广东 深圳"
                # ID|昵称|年龄|星座|地址|照片|身高|年收入|职业|恋爱次数|家乡|学校|微信
                girl_info = []
                girl_info.append(fuid)
                girl_info.append(resp["data"]["nickname"])
                girl_info.append(resp["data"]["age"])
                girl_info.append(resp["data"]["constellation"])
                girl_info.append(resp["data"]["address"])
                girl_info.append(resp["data"]["privacy"]["life_photo"]["data"])
                girl_info.append(resp["data"]["privacy"]["privacy_info"]["data"][0]["data"])
                girl_info.append(resp["data"]["privacy"]["privacy_info"]["data"][1]["data"])
                girl_info.append(resp["data"]["privacy"]["privacy_info"]["data"][2]["data"])
                girl_info.append(resp["data"]["privacy"]["privacy_info"]["data"][3]["data"])
                girl_info.append(resp["data"]["basic_info_list"]["des_info_map"]["hometown"]["info"])
                girl_info.append(resp["data"]["basic_info_list"]["des_info_map"]["school"]["info"])
                girl_info.append(resp["data"]["unlock_wechat_data"]["wechat"])

                girl_list.append(girl_info)
                if len(girl_list) == rows:
                    write_db(girl_list)
                    girl_list = []
        except Exception as e:
            print("fuid=" + str(fuid) + ",Exception:", e)
            e_fuid.append(fuid)
    if girl_list:
        write_db(girl_list)


# 爬取女嘉宾信息并记录excel文档
def get_girls_to_excel(token, echo, rows, num, fuid_begin, fuid_end, girl_list):
    for fuid in reversed(range(fuid_begin, fuid_end)):
        try:
            if fuid % echo == 0:
                print("====step:" + str(fuid) + "=====")

            # 获取用户信息
            resp = get_user_info(token, fuid)

            if resp["data"] and resp["data"]["sex_des"] == "女":
                # 将用户信息放入list，每满rows个写一次表
                resp["data"]["fuid"] = fuid
                girl_list.append(resp)
                if len(girl_list) == rows:
                    write_excel(girl_list, num)
                    girl_list = []
                    num = num + 1
                if resp["data"]["address"] == "广东 深圳":  # 可以考虑加个年龄限制，26岁以下
                    print("开始向" + str(fuid) + "号女嘉宾发起心动......")
                    # 发起心动
                    result = heart_user(token, fuid)
                    if result["message"]:
                        print("心动结果:" + result["message"])
        except Exception as e:
            print("fuid=" + str(fuid) + ",Exception:", e)


def write_db(param_list):
    # 数据库连接
    conn = pymysql.connect(host='localhost',
                           user='root',
                           password='root',
                           db='yizhoucp',
                           charset='utf8')
    # 批量插入数据
    with conn.cursor() as cursor:
        try:
            sql = "insert into custinfo values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            cursor.executemany(sql, param_list)
            conn.commit()
            print(threading.current_thread().name + ":落表成功,"+str(param_list[0][0])+"->"+str(param_list[-1][0]))
        except Exception as e:
            print(threading.current_thread().name + ":落表失败," + str(param_list[0][0]) + "->" + str(param_list[-1][0]) + ",Exception:", e)
            conn.rollback()
    conn.close()


def get_user_info(token, fuid):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Mobile Safari/537.36',
        'Connection': 'close',
        'Token': token
    }
    get_user_param = {
        'fuid': fuid,
        'from': 'recommend'
    }
    get_user_url = "https://w.yizhoucp.cn/api/apps/wcp/user/get-user-profile-start"
    requests.adapters.DEFAULT_RETRIES = 5
    response = requests.get(url=get_user_url, params=get_user_param, headers=headers)
    return response.json()


def heart_user(token, fuid):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Mobile Safari/537.36',
        'Token': token
    }
    data = {
        'form_id': (None, 'undefined'),
        'fuid': (None, fuid),
        'from': (None, 'recommend')
    }
    heart_user_url = "https://w.yizhoucp.cn/api/apps/wcp/like/heartbeat-user"
    response = requests.post(url=heart_user_url, headers=headers, files=data)
    return response.json()


def write_excel(data_list, num):
    # 创建excel工作表
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('girls')
    # 设置表头
    worksheet.write(0, 0, label='ID')
    worksheet.write(0, 1, label='昵称')
    worksheet.write(0, 2, label='年龄')
    worksheet.write(0, 3, label='星座')
    worksheet.write(0, 4, label='地址')
    worksheet.write(0, 5, label='照片')
    worksheet.write(0, 6, label='身高')
    worksheet.write(0, 7, label='年收入')
    worksheet.write(0, 8, label='职业')
    worksheet.write(0, 9, label='恋爱次数')
    worksheet.write(0, 10, label='家乡')
    worksheet.write(0, 11, label='学校')
    worksheet.write(0, 12, label='微信')
    for row in range(0, len(data_list)):
        worksheet.write(row + 1, 0, data_list[row]["data"]["fuid"])
        worksheet.write(row + 1, 1, data_list[row]["data"]["nickname"])
        worksheet.write(row + 1, 2, data_list[row]["data"]["age"])
        worksheet.write(row + 1, 3, data_list[row]["data"]["constellation"])
        worksheet.write(row + 1, 4, data_list[row]["data"]["address"])
        worksheet.write(row + 1, 5, data_list[row]["data"]["privacy"]["life_photo"]["data"])
        worksheet.write(row + 1, 6, data_list[row]["data"]["privacy"]["privacy_info"]["data"][0]["data"])
        worksheet.write(row + 1, 7, data_list[row]["data"]["privacy"]["privacy_info"]["data"][1]["data"])
        worksheet.write(row + 1, 8, data_list[row]["data"]["privacy"]["privacy_info"]["data"][2]["data"])
        worksheet.write(row + 1, 9, data_list[row]["data"]["privacy"]["privacy_info"]["data"][3]["data"])
        worksheet.write(row + 1, 10, data_list[row]["data"]["basic_info_list"]["des_info_map"]["hometown"]["info"])
        worksheet.write(row + 1, 11, data_list[row]["data"]["basic_info_list"]["des_info_map"]["school"]["info"])
        worksheet.write(row + 1, 12, data_list[row]["data"]["unlock_wechat_data"]["wechat"])
    # 保存并关闭
    workbook.save('excel\FemaleGuestInfo' + str(num) + '.xls')
    print("落表成功，序号:" + str(num))


if __name__ == "__main__":
    main()
