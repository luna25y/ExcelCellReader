#!/usr/bin/env python3

import xlwings as xw
from gtts import gTTS
import os
import tkinter as tk
from tkinter import scrolledtext
from pynput import mouse, keyboard
from threading import Thread

key_listener = None
mouse_listener = None

def read_english_word():
    if not is_playing:
        return
    
    app = xw.apps.active
    wb = app.books.active.sheets.active
    col_index = app.selection.column
    row_index = app.selection.row

    # 如果选中的是 E 列，则读取对应行 A 列的英语单词
    if col_index == 5:  # 5 对应 E 列的列索引
        english_word = wb.range(f'A{row_index}').value

        # 更新 GUI 窗口
        update_gui(english_word)

        tts = gTTS(text=english_word, lang='en')
        tts.save("temp.mp3")
        os.system("afplay temp.mp3")
        os.remove("temp.mp3")

def update_gui(english_word):
    # 在 GUI 窗口中显示英语单词
    text_widget.insert(tk.END, f'{english_word}\n')
    text_widget.see(tk.END)  # 滚动到最后一行

def play_pause():
    global is_playing
    is_playing = not is_playing
    if is_playing:
        # 启动播放
        play_button.config(text="Pause")
        Thread(target=start_listening).start()
    else:
        # 暂停播放
        play_button.config(text="Play")
        stop_listening()

def on_press(key):
    read_english_word()

def on_click(x, y, button, pressed):
    read_english_word()

def start_listening():
    # 启动监听键盘和鼠标事件
    global key_listener, mouse_listener  # 添加全局变量
    key_listener = keyboard.Listener(on_press=on_press)
    mouse_listener = mouse.Listener(on_click=on_click)
    
    key_listener.start()
    mouse_listener.start()

def stop_listening():
    # 停止监听键盘和鼠标事件
    global key_listener, mouse_listener
    if key_listener:
        key_listener.stop()
    if mouse_listener:
        mouse_listener.stop()

# GUI
root = tk.Tk()
root.title("ExcelReader")
text_widget = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=30, height=10)
text_widget.pack(expand=True, fill="both")

is_playing = False  # 初始状态为未播放

# 播放按钮
play_button = tk.Button(root, text="Play", command=play_pause)
play_button.pack()

# 启动应用程序窗口
root.mainloop()
