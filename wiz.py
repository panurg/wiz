# wiz
# v0.5
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

bp_needed = -1
selected_specie = -1

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

def calculate_species():
  human_m_x = max / 2 - max / width * (race_button_width + race_split_interval) / 2 * 3
  human_m_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2 * 3)

  elf_m_x = max / 2 - max / width * (race_button_width + race_split_interval) / 2
  elf_m_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2 * 3)

  dwarf_m_x = max / 2 + max / width * (race_button_width + race_split_interval) / 2
  dwarf_m_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2 * 3)

  porkul_m_x = max / 2 + max / width * (race_button_width + race_split_interval) / 2 * 3
  porkul_m_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2 * 3)

  human_f_x = max / 2 - max / width * (race_button_width + race_split_interval) / 2 * 3
  human_f_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2)

  elf_f_x = max / 2 - max / width * (race_button_width + race_split_interval) / 2
  elf_f_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2)

  gnome_f_x = max / 2 + max / width * (race_button_width + race_split_interval) / 2
  gnome_f_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2)

  porkul_f_x = max / 2 + max / width * (race_button_width + race_split_interval) / 2 * 3
  porkul_f_y = max - max / height * (bottom_margin + small_button_height + bottom_split_interval + race_button_height / 2)

  return [(human_m_x, human_m_y),
          (elf_m_x, elf_m_y),
          (dwarf_m_x, dwarf_m_y),
          (porkul_m_x, porkul_m_y),
          (human_f_x, human_f_y),
          (elf_f_x, elf_f_y),
          (gnome_f_x, gnome_f_y),
          (porkul_f_x, porkul_f_y)]

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

def click_specie(specie):
  win32api.mouse_event(0x8001, specie[0], specie[1])
  press()

def click_cancel():
  cancel_x = max / 2 + max / width * (small_button_width + interval_between_small_buttons) / 2
  cancel_y = max - max / height * (bottom_margin + small_button_height / 2)

  win32api.mouse_event(0x8001, cancel_x, cancel_y)
  press()

def click_ok():
  ok_x = max / 2 - max / width * (small_button_width + interval_between_small_buttons) / 2
  ok_y = max - max / height * (bottom_margin + small_button_height / 2)

  win32api.mouse_event(0x8001, ok_x, ok_y)
  press()

def click_center_ok():
  center_ok_x = max / 2
  center_ok_y = max / 2 + max / height * center_ok_shift

  win32api.mouse_event(0x8001, center_ok_x, center_ok_y)
  press()

def roll():
  box = (width / 2 - font_size, height / 2 - font_size, width / 2 + font_size, height / 2 + font_size)
  api = tesseract.TessBaseAPI()
  api.Init(".","eng",tesseract.OEM_DEFAULT)
  api.SetVariable("tessedit_char_whitelist", "0123456789")
  api.SetPageSegMode(tesseract.PSM_SINGLE_WORD)
  species = calculate_species()

  while 1:
    bonus_points = 0
    count = 0

    start.wait()
    start.clear()

    while (bonus_points < bp_needed) and (not stop.is_set()):
      click_cancel()
      click_specie(species[selected_specie])
      click_ok()
      time.sleep(2.3)
      im = ImageGrab.grab(box)
      cimg = cv.CreateImageHeader(im.size, cv.IPL_DEPTH_8U, 3)
      cv.SetData(cimg, im.tostring())
      tesseract.SetCvImage(cimg, api)
      result = api.GetUTF8Text()
      try:
        bonus_points = int(result)
      except:
        bonus_points = 0
      click_center_ok()
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

while (bp_needed < 0) or (bp_needed > 70):
  try:
    bp_needed = input("enter the desired amount of Bonus Points (min - 0, max - 70): ")
  except:
    print "NaN"
while (selected_specie < 0) or (selected_specie > 7):
  print "available species:"
  print "\t0 - Human  (M)"
  print "\t1 - Elf    (M)"
  print "\t2 - Dwarf  (M)"
  print "\t3 - Porkul (M)"
  print "\t4 - Human  (F)"
  print "\t5 - Elf    (F)"
  print "\t6 - Gnome  (F)"
  print "\t7 - Porkul (F)"
  try:
    selected_specie = input("select: ")
  except:
    print "NaN"

hm = pyHook.HookManager()
hm.KeyDown = on_keyboard_event
hm.HookKeyboard()

print "switch to Wizardry Online and press <F1> to start or <Esc> to stop"

pythoncom.PumpMessages()

rolling_thread.join()

