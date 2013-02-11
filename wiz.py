# wiz
# v0.2
# by panurg
#
# i hope u will enjoy this script :3

import win32api
import time
import ImageGrab
import tesseract
import cv2.cv as cv

bp_needed = 40

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

time.sleep(5)

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

bonus_points = 0
count = 0
box = (width / 2 - font_size, height / 2 - font_size, width / 2 + font_size, height / 2 + font_size)
api = tesseract.TessBaseAPI()
api.Init(".","eng",tesseract.OEM_DEFAULT)
api.SetVariable("tessedit_char_whitelist", "0123456789")
api.SetPageSegMode(tesseract.PSM_SINGLE_WORD)

while bonus_points < bp_needed:
  select_cancel()
  select_gnome()
  select_ok()
  time.sleep(2.2)
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
  print "rolled points = ", bonus_points
  count += 1

print "Gratz!"
print "number of attempts: ", count
