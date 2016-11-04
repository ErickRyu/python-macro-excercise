# -*- coding: utf-8 -*-

import win32gui
import re
import autopy
import time
import threading


class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""
    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name = None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        '''Pass to win32gui.EnumWindows() to check all the opened windows'''
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) != None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)

w = WindowMgr()

def find_kakao():
    w.find_window_wildcard(".*kaoTalk.*")    
    w.set_foreground()
    autopy.key.tap('l', autopy.key.MOD_CONTROL)   
    # time.sleep(0.5) 
    password = "kakaotalk password here"
    autopy.key.type_string(password)
    enter()
    time.sleep(0.5)

def find_excel():
    w.find_window_wildcard(".*Excel.*")    
    w.set_foreground()

def close_window():
    autopy.key.tap('w', autopy.key.MOD_CONTROL)

def copy_mother():
    autopy.key.tap('d', autopy.key.MOD_CONTROL)    

def day_copy():
    autopy.key.tap('t', autopy.key.MOD_CONTROL)    

def range_copy():
    autopy.key.tap('e', autopy.key.MOD_CONTROL)    

def hide_student():
    autopy.key.tap('r', autopy.key.MOD_CONTROL)    

def paste():
    autopy.key.tap('v', autopy.key.MOD_CONTROL)    

def select_all_and_delete():
    autopy.key.tap('a', autopy.key.MOD_CONTROL)    
    autopy.key.tap(autopy.key.K_DELETE)    

def enter():
    autopy.key.tap(autopy.key.K_RETURN)    

def up():
    autopy.key.tap(autopy.key.K_UP)    

def down():
    autopy.key.tap(autopy.key.K_DOWN)    

def backspace():
    autopy.key.tap(autopy.key.K_BACKSPACE)    

def wait():
    time.sleep(0.5)

if __name__ =="__main__":
    # 어머니 이름 찾아서 카톡 창에서 찾아놓기
    find_excel()
    copy_mother()
    wait()
    find_kakao()
    select_all_and_delete()
    paste()

    # 학생의 요일 정보 복사
    find_excel()
    day_copy()

    # 어머니에게 텍스트로 전송 후 어머니 카톡 창 닫음
    wait()
    find_kakao()
    enter()
    paste()
    up()
    enter()
    backspace()
    wait()
    wait()
    enter()
    close_window()
    wait()

    # 학생 점수 정보 복사
    find_excel()
    range_copy()

    # 다시 어머니에게 이미지를 전송
    wait()
    find_kakao()
    enter()
    paste()
    down()
    enter()
    enter()
    wait()
    wait()
    wait()
    close_window()

    # 학생 숨김.
    wait()
    find_excel()
    wait()
    hide_student()