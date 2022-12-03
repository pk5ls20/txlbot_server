import logging

import cv2
import numpy as np
import pyautogui
import matplotlib.pyplot as plt

#找图 返回最近似的点
def search_returnPoint(img,template,template_size):
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    template_ = cv2.cvtColor(template,cv2.COLOR_BGR2GRAY)
    result = cv2.matchTemplate(img_gray, template_,cv2.TM_CCOEFF_NORMED)
    threshold = 0.7
    # res大于70%
    loc = np.where(result >= threshold)
    # 使用灰度图像中的坐标对原始RGB图像进行标记
    point = ()
    for pt in zip(*loc[::-1]):
        cv2.rectangle(img, pt, (pt[0] + template_size[1], pt[1] + + template_size[0]), (7, 249, 151), 2)
        point = pt
    if point==():
        return -1
    return 1

def isexist():
    logging.debug("开始判断登陆状态")
    scale = 1
    pyautogui.screenshot('screen.png')
    logging.debug("oh,屏截完了")
    img = cv2.imread('screen.png')  # 要找的大图
    img = cv2.resize(img, (0, 0), fx=scale, fy=scale)
    template = cv2.imread('status.png')  # 图中的小图
    template = cv2.resize(template, (0, 0), fx=scale, fy=scale)
    template_size = template.shape[:2]
    if(search_returnPoint(img, template, template_size)==-1):
        # print("没找到图片")
        return -1
    else:
        # print("找到图片 位置:" )
        # plt.figure()
        # plt.imshow(img, animated=True)
        # plt.show()
        return 1
"""

if(img is None):
    print("没找到图片")
else:
    print("找到图片 位置:"+str(x_)+" " +str(y_))
    # plt.figure()
    # plt.imshow(img, animated=True)
    # plt.show()
"""