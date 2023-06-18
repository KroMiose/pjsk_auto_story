import datetime, os, time
import win32gui,  win32con, win32api
import pyautogui
import keyboard
import traceback

import config


# 绝对坐标点击
def clickAbsolute(x, y):
    win32api.SetCursorPos([x, y])
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

# 检测图片是否存在
def checkImgExist(src, region=None, confidence=config.MATCH_THRESHOLD):
    ''' 区域内使用pyautogui检测图片是否存在
    该方法默认是全图截屏，性能消耗较大，传入region指定区域

    参数：
    src: 图片资源路径
    region: 检测区域(x, y, width, height)
    confidence: 图片确信度（拟合程度阈值，默认0.9）

    成功返回：中心坐标(x, y)
    失败返回: False
    '''

    location = pyautogui.locateCenterOnScreen(src, region=region, confidence=confidence) if region \
        else   pyautogui.locateCenterOnScreen(src, confidence=confidence)
    if location:
        loc_x, loc_y = location
        return (loc_x, loc_y)
    else:
        return False

# 检测图片并点击
def checkImgAndClick(src, region=None, confidence=config.MATCH_THRESHOLD, clickOffset=(0, 0), afterDelay=0.5):
    ''' 使用pyautogui检测如果存在图片则点击
    该方法默认是全图截屏，性能消耗较大，传入region可指定区域

    参数：
    src: 图片资源路径
    region: 检测区域(x, y, width, height)
    confidence: 图片确信度（拟合程度阈值，默认0.9）
    clickOffset: 点击位置偏移(x, y) 以图片中心为基准
    afterDelay: 操作结束延迟，单位：秒

    成功返回：点击坐标(x, y)
    失败返回: False
    '''

    location = pyautogui.locateCenterOnScreen(src, region=region, confidence=confidence) if region \
        else   pyautogui.locateCenterOnScreen(src, confidence=confidence)
    if location:
        loc_x, loc_y = location
        x, y = loc_x + clickOffset[0], loc_y + clickOffset[1]

        log('点击了: ' + src + ' 位于: ' + str((x, y)))

        curCPos = win32api.GetCursorPos()   # 获取当前鼠标位置
        win32api.SetCursorPos([x, y])
        time.sleep(0.1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.SetCursorPos(curCPos)  # 恢复鼠标位置

        if afterDelay:
            time.sleep(afterDelay)
        # print('clicked: ', x, y)
        return (x, y)
    else:
        return False

# 获取像素颜色
def getImagePixel(img, loc):
    (b, g, r) = img[loc[1], loc[0]]
    return (r, g, b)

# 比较两个像素的颜色
def cmpPixelColor(c1, c2, offset = 3):
    if((abs(c1[0] - c2[0]) < offset) and (abs(c1[1] - c2[1]) < offset) and (abs(c1[2] - c2[2]) < offset)):
        return True
    return False

# 日志输出
def log(text):
    print('[{}] {}'.format(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S') , text))

auto_click_images = []  # 自动点击图片列表
start_cold_time = 0     # 开始章节冷却时间
finish_cold_time = 0    # 结束章节冷却时间


if __name__ == '__main__':
    print(
        f'程序: pjsk自动学习剧情脚本 {config.VERSION}\n'
        '开始程序前请保证您将游戏界面通过任意方式投屏到电脑，并且支持鼠标点击操作游戏界面\n\n'
        '即将开始进行剧情学习，请保持游戏窗口位于前台可见\n'
        '如果运行过程中出现错误或运行不正确，请编辑config.py并按要求修改配置信息或尝试更换 images 文件夹中的图片资源以提高检测准确度！\n\n'
        '提示: 本程序完全开源免费，如果你通过任何渠道购买了本程序，请立即向购买渠道索要退款！\n'
        'by: KroMiose谬锶\n\n'
    )

    # 读取所有需要自动点击的图片
    for file in os.listdir('./images/auto'):
        if(file.endswith('.png')):
            auto_click_images.append('./images/auto/' + file)
    log('读取到自动点击图片列表: {}\n'.format(auto_click_images))
    log('请手动切换到游戏窗口，然后按下F10键，程序将自动获取窗口句柄\n')
    keyboard.wait('F10')
    window_hwnd = win32gui.GetForegroundWindow()
    try:    # 调整窗口大小
        win32gui.SetWindowPos(int(window_hwnd, 16), win32con.HWND_TOP, 0, 0, config.WINDOW_WIDTH, config.WINDOW_HEIGHT, win32con.SWP_SHOWWINDOW)
    except Exception as e:
        traceback.print_exc()
        log('调整窗口时发生错误，请检查窗口句柄是否正确！')
        input('按下任意键退出程序...')
        exit(1)
    log('获取到窗口句柄: ' + window_hwnd + ' 窗口大小: ' + str(config.WINDOW_WIDTH) + 'x' + str(config.WINDOW_HEIGHT))

    while True:    # 主运行循环
        loop_sta = time.time()
        window_rect = win32gui.GetWindowRect(int(window_hwnd, 16))  # 获取窗口位置
        try:
            if time.time() - finish_cold_time > 5: # 5秒内没有点击过结束播放，自动点击结束播放
                res = checkImgAndClick(config.IMG_SRC['white_ok'], region=window_rect, confidence=0.8) # 自动勾选连续播放复选框
                if res and config.AUTO_SWITCH_CHAPTER:
                    log('检测到章节结束，自动切换章节中...')
                    for pos in config.AUTO_SWITCH_CHAPTER_COORDS:
                        clickAbsolute(window_rect[0] + pos[0], window_rect[1] + pos[1])
                        time.sleep(1.5)
                    finish_cold_time = time.time()
                    continue

            if time.time() - start_cold_time > 5: # 5秒内没有点击过开始播放，自动点击开始播放
                res = checkImgAndClick(config.IMG_SRC['continuously_check'], region=window_rect, clickOffset=(-40, 0), confidence=0.92) # 自动勾选连续播放复选框
                if res:
                    start_cold_time = time.time()
                    log('自动勾选连续播放复选框并点击开始播放')
                    clickAbsolute(res[0] + 90, res[1] + 80)
                    continue

            for img in auto_click_images:    # 匹配需要自动点击的图片
                checkImgAndClick(img, region=window_rect) # 在指定区域内自动点击图片

            log('循环监测耗时: {}s'.format(time.time() - loop_sta))
            time.sleep(config.LOOP_DELAY)

        except Exception as e:
            log(f'检测到执行异常: {e}')
            log('请尝试手动打开窗口再试！')
            time.sleep(config.LOOP_DELAY * 4)
            continue
