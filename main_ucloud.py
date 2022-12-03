# original: skywoodsz
# fix: pk5ls20
import atexit
import json
import logging
import os
import sys
import img_term
import pandas as pd
import pyautogui
import time
from datetime import datetime
import cv2
import psutil
import requests

import picfind
from picfind import isexist
from picshow import showpic

global jpath

__stderr__ = sys.stderr
current_logname = str(time.strftime("%Y-%m-%d_%H-%M-%S", time.localtime()))
sys.stderr = open(current_logname + str('.log'), 'a')


def gettime():
    return current_logname + str('.log')


logging.basicConfig(filename=gettime(), filemode="w",
                    format="%(asctime)s %(name)s:%(levelname)s:%(message)s",
                    datefmt="%d-%M-%Y %H:%M:%S", level=logging.DEBUG)

requests.packages.urllib3.disable_warnings()

with open('dconfig.json', 'r') as cf:
    dconfig = json.load(cf)
class nclass:
    classname = []
    no = []
    status = []

'''
	利用opencv模板匹配进行入会像素坐标获取，执行鼠标相应动作与密码验证存在判断
	@param tempFile: 模板匹配图像
		   whatDo：鼠标执行动作
		   debug： debug
'''


def push(title, desp):
    logging.debug("push-a-message" + str(title) + "/" + str(desp))
    api = dconfig['pushapi']
    data = {
        "title": title,
        "desp": desp}
    requests.post(api, data=data, verify=False)


def proc_exist(process_name):
    pl = psutil.pids()
    for pid in pl:
        if psutil.Process(pid).name() == process_name:
            logging.debug("pid=" + str(pid) + "killed")
            return pid


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    logging.debug('resource_path:' + str(os.path.join(base_path, relative_path)))
    return os.path.join(base_path, relative_path)


def ImgAutoClick(tempFile, whatDo, debug=False):
    pyautogui.screenshot('screen.png')
    gray = cv2.imread('screen.png', 0)
    img_templete = cv2.imread(tempFile, 0)
    w, h = img_templete.shape[::-1]
    res = cv2.matchTemplate(gray, img_templete, cv2.TM_SQDIFF)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    top = min_loc[0]
    left = min_loc[1]
    x = [top, left, w, h]
    top_left = min_loc
    bottom_right = (top_left[0] + w, top_left[1] + h)
    pyautogui.moveTo(top + h / 2, left + w / 2)

    if debug:
        print("?")
        img = cv2.imread("screen.png", 1)
        cv2.rectangle(img, top_left, bottom_right, (0, 0, 255), 2)
        img = cv2.resize(img, (0, 0), fx=0.5, fy=0.5, interpolation=cv2.INTER_NEAREST)
        cv2.imshow("processed", img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
        os.remove("screen.png")

    whatDo(x)

    if debug:
        img = cv2.imread("screen.png", 1)
        cv2.rectangle(img, top_left, bottom_right, (0, 0, 255), 2)
        img = cv2.resize(img, (0, 0), fx=0.5, fy=0.5, interpolation=cv2.INTER_NEAREST)
        cv2.imshow("processed", img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
    os.remove("screen.png")

    return True


'''
	会议自动登录
	@param meeting_id: 会议号
		   password：密码，若则保持默认NULL
'''


def SignIn(meeting_id, password=None):
    # 指定读取某一单元格内容
    sf = dconfig['tmpath']
    logging.info("当前腾讯会议路径为" + str(sf))
    # 错误indexerror
    # 先看看当前有没有腾讯会议进程，有先kill掉
    if isinstance(proc_exist('wemeetapp.exe'), int):
        os.system('taskkill /IM wemeetapp.exe /F')
        logging.info("腾讯会议已经在进行，先结束进程")
        print("腾讯会议已经在进行，先结束进程")
        time.sleep(5)
    os.startfile(sf)
    logging.info("当前腾讯会议路径为" + str(sf))
    print("当前腾讯会议路径为" + sf)
    logging.info("Start Sign, please waiting.")
    print("Start Sign, please waiting.")
    time.sleep(12)
    ImgAutoClick("JoinMeeting.png", pyautogui.click, False)
    logging.debug("clicked the JoinMeeting.png")
    time.sleep(2)
    ImgAutoClick("meeting_id.png", pyautogui.click, False)
    logging.debug("clicked the meeting_id.png")
    time.sleep(3)
    pyautogui.write(meeting_id)
    logging.debug("已注入会议号")
    time.sleep(3)
    ImgAutoClick("final.png", pyautogui.click, False)
    logging.debug("已点击进入会议")
    time.sleep(2)
    # res = ImgAutoClick("password.png", pyautogui.moveTo, False)
    if password != "xxxxxx":
        res = ImgAutoClick("password.png", pyautogui.moveTo, False)
        pyautogui.write(password)
        logging.debug("已输入参会密码")
        time.sleep(2)
        ImgAutoClick("passwordJoin.png", pyautogui.click, False)
        logging.debug("已点击进入会议")
        time.sleep(3)
    return True


def load_schdule(class_address):
    data = pd.read_excel(class_address)
    data.sort_values(by=['day', 'start_time'], inplace=True)
    data = data.reset_index(drop=True)
    # pd.set_option('max_rows', None)
    # pd.set_option('max_columns', None)
    pd.set_option('expand_frame_repr', False)
    pd.set_option('display.unicode.east_asian_width', True)
    logging.info("load_schdule")
    logging.info(str(data))
    nclass.classmate = []
    nclass.no = []
    nclass.status = []
    # 在此处建立一个字典，用于记录该课程是否成功进入过
    for i in range(len(data.lesson)):
        nclass.classname.append(data.lesson[i])
        nclass.no.append(i)
        nclass.status.append('False')
    print(data, "\n")
    return data


def get_class(schdule):
    NowDay = datetime.now().weekday()
    today_lessons = schdule.loc[schdule['day'] == NowDay + 1]
    today_lessons = today_lessons.reset_index(drop=True)
    now = datetime.now()
    for i in range(len(today_lessons)):
        logging.debug(str(today_lessons.values[i][1].hour))
        logging.debug(str(today_lessons.values[i][1].minute))
        logging.debug(str(today_lessons.values[i][1].second))
        logging.debug(str(now.hour))
        logging.debug(str(now.minute))
        Hour = today_lessons.values[i][1].hour
        Minute = today_lessons.values[i][1].minute
        Second = today_lessons.values[i][1].second
        # 在此处进行更改，确保可以延迟入会1.5h（逆天
        dtime_all = int(Hour) * 60 + int(Minute)
        ntime_all = int(now.hour) * 60 + int(now.minute)
        # 根据原逻辑这块还要重写一下，避免重复进入课堂
        if(nclass.status[int(nclass.no[i])]=='True'):
            continue
        # if now.hour < Hour or (now.hour == Hour and now.minute <= Minute):
        if ntime_all - dtime_all <= 90:
            res_time = (Hour - now.hour) * 3600 + (Minute - now.minute) * 60 + Second - now.second
            # res_time = (today_lessons[i][1].hour - now.hour)*3600 + (today_lessons[i][1].minute - now.minute)*60 - now.second
            # print(type(today_lessons.values[i][1].hour), today_lessons.values[i][1].minute, now.hour, now.minute)
            class_name, start_time, meet_id, password = today_lessons.values[i][0], today_lessons.values[i][1], \
                                                        today_lessons.values[i][2], today_lessons.values[i][3]
            print("Next Class: " + str(class_name) + "\n")
            print("Start time: " + str(start_time) + "\n")
            print("Meeting id: " + str(meet_id) + "\n")
            print("Password: " + str(password) + "\n")
            logging.info("Next Class: " + str(class_name))
            logging.info("Start time: " + str(start_time))
            logging.info("Meeting id: " + str(meet_id))
            logging.info("Password: " + str(password))
            push('听课酱要听' + str(class_name) + '啦！',
                 "Next Class: " + str(class_name) + "Start time: " + str(start_time)
                 + "Meeting id: " + str(meet_id) + "Password: " + str(password))
            while res_time > 60:
                print(
                    "Sleeping, waiting for next lesson. Res time: {} hour, {} min, {} seconds".format(res_time // 3600,
                                                                                                      (
                                                                                                                  res_time % 3600) // 60,
                                                                                                      res_time % 60))
                logging.info(
                    "Sleeping, waiting for next lesson. Res time: {} hour, {} min, {} seconds".format(res_time // 3600,
                                                                                                      (
                                                                                                                  res_time % 3600) // 60,
                                                                                                      res_time % 60))
                now = datetime.now()
                print("Now time: {} hour, {} min, {} seconds".format(now.hour, now.minute, now.second))
                logging.info("Now time: {} hour, {} min, {} seconds".format(now.hour, now.minute, now.second))
                if res_time > 3600:
                    logging.debug("Now sleeping 1h")
                    time.sleep(3600)
                    res_time -= 3600
                else:
                    time.sleep(res_time - 30)
                    logging.info("Start trying to join the class")
                    print("Start trying to join the class")
                    break
            # return lessons_name, lessons_start_time, lessons_meet_id, lessons_password
            nclass.status[int(nclass.no[i])] = 'True'
            return today_lessons.values[i][0], today_lessons.values[i][1], today_lessons.values[i][2], \
                   today_lessons.values[i][3]
    return None, None, None, None


if __name__ == "__main__":
    try:
        showpic(dconfig['moepath'])
        jpath = str(str(sys.argv[0]))
        print("---------------------------------"
              "V1.2开摆版"
              "---------------------------------")
        print("-=TxlistenbotMade=- by love with\nOriginal @skywoodsz\nFix @pk5")
        logging.info("听课酱已被唤醒~")
        push("✔听课酱已被唤醒~", '听课酱于' + jpath + '被唤醒~')
        time.sleep(9)
        print("======腾讯会议听课酱Ucloud版=======")
        print("https://github.com/pk5ls20/txlbot_server")
        print("使用前请认真阅读Readme！")
        print("================================")
        class_address = dconfig['xmlpath']
        logging.debug("class_address=" + str(class_address))
        schdule = load_schdule(class_address)
        while True:
            class_name, start_time, meet_id, password = get_class(schdule)
            logging.info("目前加载的课堂信息")
            logging.info('课程名 ' + str(class_name))
            logging.info('开始时间 ' + str(start_time))
            logging.info("会议号 " + str(meet_id))
            logging.info("会议密码 " + str(password))
            if class_name == None:
                push("听课酱说，今天剩下的时间里没有课哦", None)
                print("听课酱说，今天剩下的时间里没有课哦")
                logging.info("听课酱说，今天剩下的时间里没有课哦")
                now = datetime.now()
                # 一觉到天明qwq
                time.sleep(86400 - now.hour * 3600 - now.minute * 60 - now.second)
                push("听课酱说，又是全新的一天呢", None)
                print("听课酱说，又是全新的一天呢")
                logging.info("听课酱说，又是全新的一天呢")
                continue
            now = datetime.now()
            for i in range(1, 100):
                if picfind.isexist() == -1:  # 会议并没有登陆
                    print("少女祈祷第" + str(i) + "次中...")
                    logging.info("少女祈祷第" + str(i) + "次中...")
                    SignIn(meet_id, str(password))
                    time.sleep(10)
                else:
                    print("会议" + class_name + "经过" + str(i) + "次祈祷后已成功登陆！")
                    push("✔" + class_name + "登陆成功！", "✔会议" + class_name +
                         "经过" + str(i) + "次祈祷后已成功登陆！")
                    break
            time.sleep(200)
    except(FileNotFoundError):
        print("听课酱未找到参数文件，请检查文件是否存在")
        push("⚠听课酱未找到参数文件，请检查文件是否存在", None)
        logging.info("FileNotFoundError-"
                     "听课酱未找到参数文件，请检查文件是否存在")
        time.sleep(114514)
    except(PermissionError):
        print("听课酱无法读取参数文件，请关闭正在打开的参数文件后重试！")
        push("⚠听课酱无法读取参数文件，请关闭正在打开的参数文件后重试！", None)
        logging.info("PermissionError-"
                     "听课酱未找到参数文件，请检查文件是否存在")
        time.sleep(114514)
