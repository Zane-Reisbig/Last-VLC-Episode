import pywinauto
import psutil
import os
import sys
import pyperclip
import time
import atexit

import win32gui
import win32process
import win32api
import win32con
import win32com.client

import logging

IS_DEV = True
DO_ITER_UPDATE = False

format = logging.Formatter("%(levelname)s | %(message)s")
LOGGER = logging.getLogger()
_logStream = logging.StreamHandler()
_logStream.setFormatter(format)
LOGGER.addHandler(_logStream)
LOGGER.setLevel("DEBUG")

LOCK_FILE_NAME = os.getcwd() + "\\.loggerlock"
CONFIG_FILE_NAME = os.getcwd() + "\\config.ini"
DEFAULT_CONFIG = str(
  "EPISODE_OUT_FOLDER=Desktop\n" +
  "WATCH_FOLDER_AMOUNT=2\n" +
  "WATCH_FOLDER_1=E:\\Local Media\n" +
  "WATCH_FOLDER_2=F:\\Local Media\n"
)

EPISODE_OUT_FOLDER = ""
WATCH_FOLDERS = []

class EpisodeHelper():
  watchFolders = None
  
  def __init__(self, _watchFolders):
    self.watchFolders = _watchFolders
  
  def tryGetNextEpisode(self, fileName) -> str | None:
    fileName = os.path.basename(fileName)
    found = False
    nextFile = None      
    for watchFolder in self.watchFolders:
      for path, _, files in os.walk(watchFolder):
        if nextFile != None: break

        for file in files:
          if found == True:
            nextFile = path + "\\" +file
            break
          
          if file == fileName:
            found = True
    
    return nextFile

# Static
class Config:

  def exists(configFilePath):
    return os.path.exists(configFilePath)
  
  def load(configFilePath, loader, stripNewLines=True) -> bool:
    lines = []
    with open(configFilePath, "r") as configFile:
      lines = configFile.readlines()
   
    if stripNewLines:
      lines = [line.replace("\n", "") for line in lines]
    
    ourDict = {}
    for item in lines:
      splitIt = item.split("=")
      ourDict.update({splitIt[0] : splitIt[1]})
    
    loader(ourDict)
    
    return True
  
  def createDefault(configFilePath, defaultConfig):
    with open(configFilePath, "w") as configFile:
      configFile.write(defaultConfig)
    

# Static
class LockFile:
  def createFileLock():
    with open(LOCK_FILE_NAME, "w") as lockfile:
      lockfile.write("PID=" + str(os.getpid()) + "\n")
    
  def checkIfLockFileExists():
    if(os.path.exists(LOCK_FILE_NAME)):
      return True

    return False

  def getOldPID():
    if(not LockFile.checkIfLockFileExists()): return
    
    with open(LOCK_FILE_NAME, "r") as lockfile:
      return int(lockfile.read().split("=")[1])

  def clearLockFile():
    os.remove(LOCK_FILE_NAME)

  def killOldInstance():
    if(not LockFile.checkIfLockFileExists()):
        return
    else:
      if LockFile.getOldPID() == os.getpid():
        return
      
    try:
      psutil.Process(LockFile.getOldPID()).terminate()
    except:
      pass




# Static
class WindowHandlers:
  _shell = win32com.client.Dispatch("WScript.Shell")

  def getVLCHandle():
    processes = psutil.process_iter()
    return [item for item in processes if item.name() == "vlc.exe"][0]

  def getCurrentVLCFile(vlc):
    return vlc.open_files()

  def getCurrentWindowHandle():
    return win32gui.GetForegroundWindow()

  def getStickyWindow():
    window = win32gui.FindWindow(None, "Sticky Notes")
    pid = win32process.GetWindowThreadProcessId(window)[1]

    if pid == None or pid == 0: 
      raise Exception("Sticky Notes not started")
    
    app:pywinauto.Application = pywinauto.Application(backend="uia").connect(process=pid)

    return app.top_window()

  def isWindowFullscreen(hwnd):
    # If pywin32 had getWindowInfo id be a happy man
    
    # This might not work :/
    windowRect = win32gui.GetClientRect(hwnd)
    if (
      windowRect[0] == 0
      and windowRect[1] == 0 
      and windowRect[2] == win32api.GetSystemMetrics(win32con.SM_CXSCREEN) 
      and windowRect[3] == win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    ): return True

    # This ALSO might not work
    if (
      # Logical Operator Nonsense
      win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE) & win32con.WS_POPUP != 0
    ): return True
    
    return False

  def createShortcut(sourcePath, placeAt="Desktop", name=None) -> str:
    linkName = os.path.basename(placeAt).split(".")[0] if name == None else name
    if placeAt == "Desktop":
      linkName = f"{WindowHandlers._shell.SpecialFolders('Desktop')}\\{linkName}.lnk"
    else:
      linkName = f"{placeAt}\\{linkName}.lnk"
      
    short = WindowHandlers._shell.CreateShortcut(linkName)
    short.TargetPath = sourcePath
    short.Save()

    return linkName
  
  def tryRemoveShortcut(sourcePath) -> bool:
    if not os.path.exists(sourcePath):
      return False
    
    os.remove(sourcePath)
    return True



# Static
class EditController:
  def _assertStickyIsSelected(target):
    target.set_focus()
  
  def clearAllText(edit):
    EditController._assertStickyIsSelected(edit)
    edit.type_keys("^a")
    EditController._assertStickyIsSelected(edit)
    edit.type_keys("{BACKSPACE}")

  def setEditText(edit, text, doPaste=False):
    EditController.clearAllText(edit)
    time.sleep(0.2)
    
    if doPaste:
      hold = pyperclip.paste()
      pyperclip.copy(text)
      EditController._assertStickyIsSelected(edit)
      edit.type_keys("^v")
      time.sleep(0.3)
      pyperclip.copy(hold)
      return
      
    EditController._assertStickyIsSelected(edit)
    edit.type_keys(text.replace("\n", "{ENTER}}"))
    
    # This son of a gun doesnt want to work well
    EditController.setEditText(edit, text, doPaste)
    EditController.setEditText(edit, text, doPaste)




def main():
  LOGGER.debug("Loading Main")
  
  epHelper = EpisodeHelper(WATCH_FOLDERS)
  sticky = WindowHandlers.getStickyWindow()
  vlc = WindowHandlers.getVLCHandle()
  lastWindowHandle = None
  outText = "[VLC LAST MEDIA]\n"
  
  lastFile = []
  
  d_iter = 1
  LOGGER.debug("Loaded Main, starting main loop...")
  while True:
    if WindowHandlers.isWindowFullscreen(WindowHandlers.getCurrentWindowHandle()):
      LOGGER.debug("Full Screen Window Detected, waiting...")
      time.sleep(60)
      continue
    
    try:
      vlc = WindowHandlers.getVLCHandle()
      sticky = WindowHandlers.getStickyWindow()
      LOGGER.debug("Got handles for iteration: %d", d_iter)
    except:
      pass
      
      # if they aint open, wait and try again
      LOGGER.debug("Failed to get Handles, waiting...")
      time.sleep(5)
      continue
      
    hold = None
    
    for cur_file in vlc.open_files():
      if hold != None: break
      
      for prefix in WATCH_FOLDERS:
        if prefix in cur_file.path:
          hold = cur_file
          break
    
    # No file, not a problem
    if hold == None:
      LOGGER.debug("Failed to get any file from vlc")
      time.sleep(5)
      continue
    
    LOGGER.debug("Got file %s", hold.path)
    if lastFile != hold or (IS_DEV and DO_ITER_UPDATE):
      curShort = WindowHandlers.createShortcut(hold.path, EPISODE_OUT_FOLDER, "Current Episode")
      LOGGER.debug("Shortcut for cur file created at %s", curShort)

      nextEp = epHelper.tryGetNextEpisode(hold.path)  
      if nextEp != None:
        nextShort = WindowHandlers.createShortcut(nextEp, EPISODE_OUT_FOLDER, "Next Episode")
        LOGGER.debug("Shortcut for next file created at %s", nextShort)

      outText += os.path.basename(hold.path)
      outText += f"\n{hold.path}"
      
      if IS_DEV and DO_ITER_UPDATE:
        outText += f"\n[DEBUG ITER {d_iter}]"
      
      lastWindowHandle = WindowHandlers.getCurrentWindowHandle()
      sticky.set_focus()
      
      EditController.setEditText(sticky, outText + "\n", doPaste=True)
      LOGGER.debug("Edit Text Set")
      
      win32gui.SetForegroundWindow(lastWindowHandle)
    else:
      # Wait for a new file
      time.sleep(5)
    
    d_iter += 1
    lastFile = hold
    outText = "[VLC LAST MEDIA]\n"
    LOGGER.debug("Loop Complete, sleeping")
    time.sleep(0.5)


def configLoadDelegate(configDict):
  global EPISODE_OUT_FOLDER, WATCH_FOLDERS
  EPISODE_OUT_FOLDER = configDict["EPISODE_OUT_FOLDER"]
  
  for n in range(int(configDict["WATCH_FOLDER_AMOUNT"])):
    WATCH_FOLDERS.append(configDict[f"WATCH_FOLDER_{n + 1}"])



if Config.exists(CONFIG_FILE_NAME):
  try:
    Config.load(CONFIG_FILE_NAME, configLoadDelegate)
  except:
    Config.createDefault(CONFIG_FILE_NAME, DEFAULT_CONFIG)
  finally:
    Config.load(CONFIG_FILE_NAME, configLoadDelegate)
else:
  Config.createDefault(CONFIG_FILE_NAME, DEFAULT_CONFIG)
  Config.load(CONFIG_FILE_NAME, configLoadDelegate)

atexit.register(LockFile.clearLockFile)
LockFile.killOldInstance()
LockFile.createFileLock()

main()
    
  
