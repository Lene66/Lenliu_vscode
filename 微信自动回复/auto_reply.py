import pyautogui
import pyperclip
import time
import xlrd

def mouseClick(img,lr,):
    while True:
        location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
        if location is not None:
            pyautogui.click(location.x/2,location.y/2,button=lr)
            break
        print('目前没有信息 10s后重试')
        time.sleep(10)

def mainWork():
    i=1
    while i<sheet1.nrows:
        cmdType=sheet1.row(i)[0]
        if cmdType.value == 1.0:
            img=sheet1.row(i)[1].value
            mouseClick(img,'left')
        elif cmdType.value == 3.0:
            img=sheet1.row(i)[1].value
            mouseClick(img,'left')
        elif cmdType.value ==4.0:
            Intputvalue=sheet1.row(i)[1].value
            pyperclip.copy(Intputvalue)
            pyautogui.hotkey('command','v')
            pyautogui.hotkey('enter')
        elif cmdType.value == 5.0:
            img=sheet1.row(i)[1].value
            mouseClick(img,'left')
        elif cmdType.value == 6.0:
            sleepTime=sheet1.row(i)[1].value
            time.sleep(sleepTime)
        i+=1


if __name__ == '__main__':
    file='cmd.xls'
    xlSx=xlrd.open_workbook(filename=file)
    sheet1=xlSx.sheet_by_name('Sheet1')
    keY=input('选择1:执行一次 选择2:一直循环\n')
    if keY=='1':
        mainWork()
    elif keY == '2':
        while True:
            mainWork()
            time.sleep(0.2)
            print('执行完成 等待0.2s...')


