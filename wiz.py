# wiz
# v0.1
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

gnome_x = max * 4 / 7
gnome_y = max * 6 / 7

cancel_x = max * 4 / 7
cancel_y = max * 19 / 20

ok_x = max * 3 / 7
ok_y = max * 19 / 20

center_ok_x = max / 2
center_ok_y = max * 11 / 20

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

#win32api.mouse_event(0x8001, gnome_x, gnome_y)
#win32api.mouse_event(0x8001, cancel_x, cancel_y)
#win32api.mouse_event(0x8001, ok_x, ok_y)
#win32api.mouse_event(0x8001, center_ok_x, center_ok_y)

bonus_points = 0
box = (490, 365, 540, 395)
api = tesseract.TessBaseAPI()
api.Init(".","eng",tesseract.OEM_DEFAULT)
api.SetVariable("tessedit_char_whitelist", "0123456789")
api.SetPageSegMode(tesseract.PSM_SINGLE_WORD)

#file = "screen.jpg"


while bonus_points < bp_needed:
  select_cancel()
  select_gnome()
  select_ok()
  time.sleep(2.2)
  im = ImageGrab.grab(box)
  #im.save(file)
  #buffer = open(file, "rb").read()
  #result = tesseract.ProcessPagesBuffer(buffer, len(buffer), api)
  cimg = cv.CreateImageHeader(im.size, cv.IPL_DEPTH_8U, 3)
  cv.SetData(cimg, im.tostring())
  tesseract.SetCvImage(cimg, api)
  result = api.GetUTF8Text()
  print "rolled points = ", result
  if result != "":
    bonus_points = int(result)
  select_center_ok()
  if bonus_points > 70:
    bonus_points = 0

print "Gratz!"
