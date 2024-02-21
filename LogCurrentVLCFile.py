import pywinauto
import win32gui
import win32process
import win32api
import win32con
import psutil
import os
import pyperclip
import time
import atexit

LOCK_FILE_NAME = os.getcwd() + "\\.loggerlock"

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


class WindowHandlers:
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


def main():
  atexit.register(LockFile.clearLockFile)
  LockFile.killOldInstance()
  LockFile.createFileLock()
  
  watchFolderPrefix = [
    "E:\\Local Media",
    "F:\\Local Media",
  ]
  
  sticky = WindowHandlers.getStickyWindow()
  vlc = WindowHandlers.getVLCHandle()
  lastWindowHandle = None
  outText = "[VLC LAST MEDIA]\n"
  
  lastFiles = []
  
  isDev = False
  d_iter = 1
  while True:
    if WindowHandlers.isWindowFullscreen(WindowHandlers.getCurrentWindowHandle()):
      time.sleep(60)
      continue
    
    try:
      vlc = WindowHandlers.getVLCHandle()
      sticky = WindowHandlers.getStickyWindow()
    except:
      pass
      
      # if they aint open, wait and try again
      time.sleep(5)
      continue
      
    hold = []
    
    for cur_file in vlc.open_files():
      hold.append(cur_file)
    
    for file in hold:
      if lastFiles != hold or isDev:
        for prefix in watchFolderPrefix:
          if prefix in file.path:
            outText += os.path.basename(file.path)
            outText += f"\n{file.path}"
            if isDev:
              outText += f"\n[DEBUG ITER {d_iter}]"
            lastWindowHandle = WindowHandlers.getCurrentWindowHandle()
            sticky.set_focus()
            EditController.setEditText(sticky, outText + "\n", doPaste=True)
            win32gui.SetForegroundWindow(lastWindowHandle)
    
    d_iter += 1
    lastFiles = hold
    outText = "[VLC LAST MEDIA]\n"
    time.sleep(0.5)
        
main()
  
