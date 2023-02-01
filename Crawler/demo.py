import time
import datetime
import pandas as pd
import xlwt

from Config import *  # 网页xpath和selector配置
from PIL import Image
from tkinter import Tk
from distutils.log import error
from pyquery import PyQuery as pq
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from tkinter.messagebox import showinfo, showwarning, showerror
from prettytable import PrettyTable


def beauty_print(dataframe, header):
    """
    美化命令端输出
    :param dataframe:数据框类型数据
    :param header:表头
    """
    table = PrettyTable()
    table.field_names = header
    for index in range(dataframe.shape[0]):
        table.add_row(dataframe.iloc[index])
    table.align = "l"
    print(table)


def beauty_export(sheet_name, dataframe, save_path):
    """
    设置输出的Excel格式
    :param sheet_name:sheet名
    :param dataframe:数据框类型数据
    :param save_path:保存路径
    """
    wb = xlwt.Workbook()
    sheet = wb.add_sheet(sheetname=sheet_name, cell_overwrite_ok=True)

    header_style = xlwt.XFStyle()
    body_style = xlwt.XFStyle()

    header_font = xlwt.Font()
    body_font = xlwt.Font()

    header_font.name = body_font.name = '微软雅黑'
    header_font.bold = True
    header_font.height = body_font.height = 200

    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_LEFT  # 水平左对齐
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 上下居中对齐

    for col in range(dataframe.shape[1]):
        sheet.col(col).width = 25 * 256

    header_style.font = header_font
    body_style.font = body_font

    header_style.alignment = body_style.alignment = alignment

    header = dataframe.columns.tolist()

    for col in range(len(header)):
        sheet.write(0, col, header[col], header_style)

    for row in range(dataframe.shape[0]):
        for col in range(dataframe.shape[1]):
            sheet.write(row + 1, col, dataframe.iloc[row, col], body_style)

    wb.save(save_path)


def get_slideImg(driver, xpath, name):
    """
    滑块验证：获取背景图和小滑块图
    :param driver: webdriver
    :param xpath: 网页xpath
    :param name: 图片保存名称
    :return: 屏幕中的位置、宽高
    """
    # 获取屏幕，进行裁剪保存滑块验证图
    img = driver.find_element(By.XPATH, xpath)
    location = img.location
    size = img.size
    left, up, right, down = location["x"], location["y"], location["x"] + size["width"], location["y"] + size["height"]
    # 保存屏幕截图
    driver.save_screenshot("bg.png")
    # 截取验证图
    need_bg_img = Image.open("bg.png")
    img_writer = need_bg_img.crop((left, up, right, down))  # 指定上下左右截取
    img_writer.save(name)
    return location, size


def get_track(distance):
    """
    模拟人为滑动 (先慢 中快 后慢)获取移动轨迹
    :param distance: 偏移量
    :return: 移动轨迹列表
    """
    track = []
    current = 0  # 当前位移
    mid = distance * 1.5  # 设定一个阈值进行改变加速度
    t = 5  # 计算间隔
    v = 0  # 初速度
    while current < distance:
        if current < mid:
            a = 1  # 加速度为正1
        else:
            a = -2  # 加速度为负2
        v0 = v  # 初速度v0
        v = v0 + a * t  # 当前速度v = v0 + at
        # 移动距离x = v0t + 1/2 * a * t^2
        move = v0 * t + 1 / 2 * a * t * t
        current += move  # 当前位移
        track.append(round(move))  # 加入轨迹
    track.append(distance - sum(track))
    return track


def _slide(driver):
    """
    滑动滑块进行验证
    方法是通过每次滑动小滑块本身2/3的距离逐步逼近，直至验证通过，不通过则刷新网页重新验证（笨方法。。。）
    :param driver: webdriver
    """
    button = driver.find_element(By.XPATH, Slide_xpath)

    slideImg_loc, slideImg_size = get_slideImg(driver, Verify_xpath, '验证图片.png')  # 获取验证图
    slideImg_block_loc, slideImg_block_size = get_slideImg(driver, Slide_xpath, '缺口图片.png')  # 获取缺口

    for i in range(round(slideImg_size["width"] / slideImg_block_size['width']) + 1):
        distance = slideImg_block_size['width'] * 2 / 3 * (i + 1)
        dis_list = get_track(distance)
        try:
            ActionChains(driver).click_and_hold(button).perform()
            for dis in dis_list:
                ActionChains(driver).move_by_offset(xoffset=dis, yoffset=0).perform()
            ActionChains(driver).release().perform()
            time.sleep(0.3)
        except:
            break
        print('...')


def load_login_page(driver, wait_time):
    input_name = wait_time.until(EC.presence_of_element_located((By.CSS_SELECTOR, User_selector)))
    input_name.send_keys(input('请输入用户名：', ))
    input_password = wait_time.until(EC.presence_of_element_located((By.CSS_SELECTOR, Password_selector)))
    input_password.send_keys(input('请输入密码：', ))

    yanZheng = wait_time.until(EC.element_to_be_clickable((By.CSS_SELECTOR, Verify_selector)))
    yanZheng.click()

    # refreshBtn
    refresh = wait_time.until(EC.element_to_be_clickable((By.CSS_SELECTOR, Refresh_selector)))
    refresh.click()

    _slide(driver)
    time.sleep(1)

    submit = wait_time.until(EC.presence_of_element_located((By.CSS_SELECTOR, Submit_selector)))
    submit.click()

    try:
        error_msg = driver.find_element(By.XPATH, Error_msg_xpath)
        print(error_msg.text)
        print('重新验证...')
        return error_msg.text
    except:
        return None


def main():
    """
    使用谷歌浏览器
    option = webdriver.ChromeOptions()
    option.add_argument("headless")
    driver = webdriver.Chrome(chrome_options=option)
    """
    # 使用火狐浏览器
    # driver = webdriver.Firefox()
    # driver.maximize_window() # 浏览器最大化 可能不生效

    # 不显示浏览器 后台操作
    option = webdriver.FirefoxOptions()
    option.add_argument("--headless")
    driver = webdriver.Firefox(options=option)

    wait_time = WebDriverWait(driver, 10)

    # tips = Tk()
    # tips.withdraw()

    driver.get(url)
    msg = load_login_page(driver, wait_time)
    if msg:
        driver.close()
        return main()
    else:
        print('已完成验证！')


if __name__ == '__main__':
    main()
