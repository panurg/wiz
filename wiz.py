# wiz
# v0.4
# by panurg
#
# i hope u will enjoy this script :3

import win32api
import time
import ImageGrab
import tesseract
import cv2.cv as cv
import pythoncom
import pyHook
import threading
import numpy.core.multiarray # workaround for Pyinstaller

bp_needed = 40

start_key = 'F1'
stop_key = 'Escape'

game_name = 'Wizardry Online'

max = 65535

width = win32api.GetSystemMetrics(0)
height = win32api.GetSystemMetrics(1)

font_size = 20

race_button_width = 147
race_button_height = 37
small_button_width = 115
small_button_height = 22
race_split_interval = 7
bottom_margin = 30
bottom_split_interval = 38
interval_between_small_buttons = 46
center_ok_shift = 37

gnome_x = max / 2 + max / width * (race_button_width + race_split_interval) / 2
gnome_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2)

cancel_x = max / 2 + max / width * (small_button_width + interval_between_small_buttons) / 2
cancel_y = max - max / height * (bottom_margin + small_button_height / 2)

ok_x = max / 2 - max / width * (small_button_width + interval_between_small_buttons) / 2
ok_y = cancel_y

center_ok_x = max / 2
center_ok_y = max / 2 + max / height * center_ok_shift

def on_keyboard_event(event):
  if event.WindowName == game_name:
    if event.Key == start_key:
      print "rolling..."
      start.set()
    if event.Key == stop_key:
      print "stopped"
      stop.set()
  return True

def press():
  delay = 0.1
  win32api.mouse_event(0x0001, 1, 1)
  time.sleep(delay)
  win32api.mouse_event(0x0002, 0, 0)
  time.sleep(delay)
  win32api.mouse_event(0x0004, 0, 0)
  time.sleep(delay)

def select_gnome():
  win32api.mouse_event(0x8001, gnome_x, gnome_y)
  press()

def select_cancel():
  win32api.mouse_event(0x8001, cancel_x, cancel_y)
  press()

def select_ok():
  win32api.mouse_event(0x8001, ok_x, ok_y)
  press()

def select_center_ok():
  win32api.mouse_event(0x8001, center_ok_x, center_ok_y)
  press()

def roll():
  box = (width / 2 - font_size, height / 2 - font_size, width / 2 + font_size, height / 2 + font_size)
  api = tesseract.TessBaseAPI()
  api.Init(".","eng",tesseract.OEM_DEFAULT)
  api.SetVariable("tessedit_char_whitelist", "0123456789")
  api.SetPageSegMode(tesseract.PSM_SINGLE_WORD)

  while 1:
    bonus_points = 0
    count = 0
    
    start.wait()
    start.clear()

    while (bonus_points < bp_needed) and (not stop.is_set()):
      select_cancel()
      select_gnome()
      select_ok()
      time.sleep(2.3)
      im = ImageGrab.grab(box)
      cimg = cv.CreateImageHeader(im.size, cv.IPL_DEPTH_8U, 3)
      cv.SetData(cimg, im.tostring())
      tesseract.SetCvImage(cimg, api)
      result = api.GetUTF8Text()
      if result != "":
        bonus_points = int(result)
      select_center_ok()
      if bonus_points > 70:
        bonus_points = 0
      print "points = ", bonus_points
      count += 1
    
    if not stop.is_set():
      print "Gratz!"
      print "number of attempts: ", count
     
    stop.clear()

start = threading.Event()
stop = threading.Event()

rolling_thread = threading.Thread(target = roll)

rolling_thread.start()

hm = pyHook.HookManager()
hm.KeyDown = on_keyboard_event
hm.HookKeyboard()

print "press <F1> to start or <Esc> to stop"

pythoncom.PumpMessages()

rolling_thread.join()
