from pywinauto import Desktop
import win32gui
import win32con
import win32api
from selenium import webdriver
from selenium.webdriver.common.by import By
import threading
import psutil
import tkinter as tk
import tkinter.ttk as ttk
import pygubu
import time

PROJECT_UI = "ui-file.ui"
WORDLE_URL = "https://www.wordleunlimited.com/"

key_map = {
    "a": 0x41,
    "b": 0x42,
    "c": 0x43,
    "d": 0x44,
    "e": 0x45,
    "f": 0x46,
    "g": 0x47,
    "h": 0x48,
    "i": 0x49,
    "j": 0x4A,
    "k": 0x4B,
    "l": 0x4C,
    "m": 0x4D,
    "n": 0x4E,
    "o": 0x4F,
    "p": 0x50,
    "q": 0x51,
    "r": 0x52,
    "s": 0x53,
    "t": 0x54,
    "u": 0x55,
    "v": 0x56,
    "w": 0x57,
    "x": 0x58,
    "y": 0x59,
    "z": 0x5A,
    "enter": win32con.VK_RETURN,
    "backspace": win32con.VK_BACK
}

def log(log_message: str, log_type: str):
    """System logging controller

    log_message -> str -- log message
    log_type -> str -- message logging level (i.e. error, warning, log)
    """
    if True:
        print(f"{log_type.upper()}: {log_message}")

class Dictionary:
    """
    """

    def __init__(self):
        """"""

    def next(self):
        pass

class Solver:
    """Represents a Wordle Solver session

    Methods
    -------
        Browser
        -------
        open_browser(self, browser: str, url: str)
            Starts browser session
        def close_browser(self)
            Closes browser session
        def get_browser_pid(self) -> int
            Gets browser sessions process id
        def set_browser_pid(self, pid: int)
            Sets browser session process id
        def get_browser_driver(self) -> webdriver.chrome.webdriver.WebDriver
            Gets browser driver
        def set_browser_driver(self, driver: webdriver.chrome.webdriver.WebDriver)
            Sets browser driver   
    """

    def __init__(self):
        """"""
        self.driver = None
        self.process_id = None
        self.paused = None
        self.run = None
        self.hwndMain = None
        self.hwndChild = None

    def open_browser(self, browser: str, url: str):
        """Starts browser session

        browser -> (str) -- specify browser type (supported: chrome, firefox)
        url -> (str) -- url of webpage
        """
        if browser == "chrome":
            self.driver = webdriver.Chrome()
            self.driver.get(url)
            self.hwndMain = win32gui.FindWindow("Chrome_WidgetWin_1", "Wordle Unlimited - Play Wordle with Unlimited Words - Google Chrome")
            self.hwndChild = win32gui.GetWindow(self.hwndMain, win32con.GW_CHILD)
        elif browser == "firefox":
            self.driver = webdriver.Firefox()
        else:
            pass

        self.process_id = psutil.Process(self.driver.service.process.pid).children()[0].pid
        log(f"{self.process_id} {browser} browser started successfully", "log")
        self.set_browser_driver(self.driver)

    def close_browser(self):
        """Closes browser session
        """
        self.driver.quit()
        log(f"{self.process_id} browser driver closed", "log")
        self.process_id = None

    def get_browser_pid(self) -> int:
        """Gets browser sessions process id

        return -> (int) -- process id
        """
        return self.process_id

    def set_browser_pid(self, pid: int):
        """Sets browser session process id

        pid -> (int) -- browser process id        
        """
        self.process_id = pid

    def get_browser_driver(self) -> webdriver.chrome.webdriver.WebDriver:
        """Gets browser driver

        return -> (webdriver.chrome.webdriver.WebDriver) -- browser session driver
        """
        return self.driver

    def set_browser_driver(self, driver: webdriver.chrome.webdriver.WebDriver):
        """Sets browser driver
        
        driver -> (webdriver.chrome.webdriver.WebDriver) -- browser session driver
        """
        self.driver = driver

    def start_solver(self, mode: str):
        """Spawns threading to run solver

        mode -> (str) -- Auto/Manual solver mode
        """
        if mode == "auto":
            self.run = True
            self.paused = False
            threading.Thread(target=self._auto_solver).start()
            log("Auto solver started", "log")
        elif mode == "manual":
            self.run = True
            self.paused = False
            threading.Thread(target=self._manual_solver).start()
            log("Manual solver started", "log")
        else:
            log("Invalid solver type", "error")

    def pause_solver(self):
        """Pause current running solver
        """
        self.paused = not(self.paused)
        log("Solver Paused", "log")

    def stop_solver(self):
        """
        """
        self.run = False
        self.paused = True
        log("Solver Stopped", "log")

    def word_list(self, word: str, enter=False) -> list:
        """
        """
        temp_list = []
        for letter in word:
            temp_list.append(letter)

        if enter:
            temp_list.append("enter")

        return temp_list

    def typing(self, keys: list):
        """
        """
        print("typing")
        win32gui.SetForegroundWindow(self.hwndMain)
        for stroke in keys:
            win32api.SendMessage(self.hwndChild, win32con.WM_KEYDOWN, key_map[f"{stroke}"], 0)

    def _auto_solver(self):
        """
        """
        while self.run:
            self.typing(self.word_list("raise", True))
            print(f"Run: {self.run} Paused: {self.paused}")
            time.sleep(2)

    def _manual_solver(self):
        """
        """
        correction_factor = 0
        while self.run:
            time.sleep(2)
            correction_factor += 2
            while not(self.paused):
                print("manual")
                time.sleep(2 + correction_factor)
                correction_factor = 0


class SolverGUI:
    """Represents a Wordle Solver GUI session

    Methods
    -------
    """

    def __init__(self, solver: Solver, master = None):
        """
        
        """
        self.builder = builder = pygubu.Builder()
        builder.add_from_file(PROJECT_UI)
        self.mainwindow = builder.get_object("mainwindow", master)
        
        self.progress_var = None
        self.speed_var = None
        builder.import_variables(self, ["progress_var", "speed_var"])
        
        builder.connect_callbacks(self)

        self.solver = solver
    
    def run(self):
        self.mainwindow.mainloop()

    def on_open_button(self):
        """
        
        """
        threading.Thread(target=self.solver.open_browser, args=("chrome", WORDLE_URL)).start()
        #self.progress_var.set(self.progress_var.get() + 1)       

    def on_start_button(self):
        """
        
        """
        threading.Thread(target=self.solver.start_solver("auto")).start()
        self.progress_var.set(self.progress_var.get() + 1)

    def on_pause_button(self):
        """
        
        """
        threading.Thread(target=self.solver.pause_solver).start()

    def on_stop_button(self):
        """
        
        """
        threading.Thread(target=self.solver.stop_solver).start()

    def on_close_button(self):
        """
        
        """
        threading.Thread(target=self.solver.close_browser).start()
        self.progress_var.set(0)

if __name__ == "__main__":
    root = tk.Tk()
    solver = Solver()
    app = SolverGUI(solver, root)
    app.run()
    
