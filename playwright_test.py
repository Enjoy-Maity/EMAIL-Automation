import os
import time
# import sqlite
from tkinter import *
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
from PIL import ImageTk, Image
from playwright.sync_api import Playwright, sync_playwright, expect
import keyboard

class Get_authenticator_code():
    def __init__(self) -> None:
        self.app = tk.Tk()
        self.app.iconbitmap("./images/ericsson-blue-icon-logo.ico")
        self.app.geometry("520x150")
        self.app.minsize(520,150)
        self.app.maxsize(520,150)
        self.app.title("    Microsoft Authenticator Code")

        self.app.bind('<Return>', self.submit)
        self.app_background_image = ImageTk.PhotoImage(Image.open("./images/MPBN PLANNING TASK_3_5.png"))
        self.app.canvas = tk.Canvas(self.app,width=520, height=150,bd=0, highlightthickness=0, relief="ridge")
        self.app.canvas.grid(row=0,column=0,sticky=NW)

        self.app.canvas.create_image(0,0, image=self.app_background_image, anchor="nw")
        self.app.canvas.create_text(260, 35, text="Please Enter the Microsoft Authenticator Code to proceed!", font= ("Ericsson Hilda",14, "bold"), fill= "#FFFFFF")

        self.code = StringVar()
        self.entry = tk.Entry(self.app, width=45, show='*', textvariable=self.code, font=("Ericsson Hilda", 12))
        self.app.canvas.create_window(218,70, window=self.entry)
        self.entry.focus_force()
        self.app.button = tk.Button(self.app, text="Submit", font=("Ericsson Hilda", 12), command= lambda : self.submit(""))
        self.app.canvas.create_window(450,120, window=self.app.button)
        # while len(self.code.get()) == 0:
        #     self.app.show()
        self.app.mainloop()

    def submit(self, _):
        if len(self.code.get()) == 0:
            messagebox.showerror(title = "Code not entered!", message= "Please! Enter the authenticator code to proceed!")

        elif any((character.isalpha()) or (character.isspace()) for character in self.code.get()):
            messagebox.showerror(title= "Wrong Input detected!", message="Letters or whitespaces have been entered! Only Numbers are allowed!")
            self.entry.delete(0, END)

        elif len(self.code.get()) < 6:
            messagebox.showerror(title= "Wrong Input Code", message="Please enter the correct and complete authentication code from your \'Microsoft Authenticator\'")
            self.entry.delete(0, END)

        else:
            self.app.withdraw()
            self.app.destroy()

    def get_code(self):
        if len(self.code.get()) > 0:
            return int(float(self.code.get()))
        else:
            return 0

def exception_raiser(exception,*args):
    if(exception == 'Exception'):
        if(len(args) == 0):
            raise Exception()

def run(playwright:Playwright) ->None:
    # browser_type = playwright.chromium
    browser = playwright.chromium.launch(executable_path=r'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', headless = False, slow_mo = 100)
    context = browser.new_context()
    context.set_default_timeout(10000)

    # Opening the page
    page = context.new_page()

    # Go to the ITSM website
    page.goto("https://nextgentm-in.sdt.ericsson.net/arsys/forms/umt-ars-in/SHR%3ALandingConsole/Default+Administrator+View/?cacheid=a9eea6e3")

    # finding the id 'domain' and 'Employee' in the option value
    select_element = page.locator("select#domain")
    select_element.select_option("Employee")
    page.wait_for_selector("//a[@id='loginBtn']")
    page.locator("//a[@id='loginBtn']",has_text="Log On").click()

    # Getting the username via subprocess 
    # cmd /c runs the command in cmd and terminate
    # cmd /k runs the command in cmd and remain
    # username = subprocess.run((r"cmd /c",r"'echo %username%'"),capture_output=True,shell=False)

    # Getting the username via os.system
    # cmd /c runs the command in cmd and terminate
    # cmd /k runs the command in cms and remain
    # username = subprocess.run((r"cmd /c",r"'echo %username%'"),capture_output=True,shell=False)
    username = os.popen(r'cmd /c "echo %username%"').read().upper()

    # Entering the username
    # page.wait_for_selector("//input[@id='login']")
    page.locator("//input[@id='login']").fill(username)

    try:
        # Entering the passwd
        page.locator("//input[@id='passwd']").fill("heqt65$$IR$$")

        # clicking the logon button
        page.locator("//a[@id='loginBtn']",has_text='Log On').click()

        # # page.wait_for_timeout(5000)
        page.wait_for_load_state()

        # print(len(page.get_by_text("Try again after some time or contact your help desk").all()))
        expect(page.get_by_text("Try again after some time or contact your help desk")).to_be_hidden()

        # # page.on("dialog", lambda dialog: dialog.dismiss())

        app = Get_authenticator_code()
        page.locator("//input[@id='response']").focus()
        page.locator("//input[@id='response']").fill(str(app.get_code()))
        page.locator("//a[@id= 'ns-dialogue-submit']", has_text='Submit').click()
        # page.wait_for_timeout(5000)

        #testing
        # page.goto("https://nextgentm-in.sdt.ericsson.net/arsys/forms/umt-ars-in/SHR%3ALandingConsole/Default+Administrator+View/?cacheid=dc6a69f2")


        # time.sleep(3)
        page.locator("//img[@alt='Show Application List']").click()
        # time.sleep(2)
        page.locator("//div[@arid='app1603']", has_text='Change Management').hover()
        time.sleep(2)
        page.locator("//a[@class='btn']", has_text='Search Change').click()
        time.sleep(2)

        page.locator("//a[@class='advancedsearch' and @arwindowid='3']", has_text="Advanced search").click()
        time.sleep(1)

        # page.mouse.wheel(0,15000)
        page.mouse.down()
        time.sleep(1)

        # filled_space = f"('Scheduled Start Date+' >= "01/15/2024 22:00:00" AND 'Scheduled End Date+' <= "01/16/2024 08:00:00") AND ('Coordinator Group*+' = "SRF-MPBN CM Delhi")"
        filled_space = f"('Scheduled Start Date+' >= \"{datetime.now().strftime('%m')}/{datetime.now().strftime('%d')}/{datetime.now().strftime('%Y')} 22:00:00\" AND 'Scheduled End Date+' <= \"{(datetime.now()+timedelta(days=1)).strftime('%m')}/{(datetime.now()+timedelta(days=1)).strftime('%d')}/{(datetime.now()+timedelta(days=1)).strftime('%Y')} 08:00:00\") AND ('Coordinator Group*+' = \"SRF-MPBN CM Delhi\")"
        page.locator("//textarea[@class='sr' and @id='arid1005']").fill(filled_space)
        time.sleep(10)

        # page.mouse.wheel(0,)
        page.mouse.up()

    except Exception as e:
        from tkinter import messagebox
        import traceback
        messagebox.showerror(e.__class__.__name__,f"{traceback.format_exc()}\n\n{e}")
        # context.wait_for_event("close", timeout=0)
        browser.close()
        playwright.stop()
        keyboard.press_and_release("ctrl+c")
        del playwright
        # playwright.dispose()

    # context.wait_for_event("close", timeout=0)
    # browser.close()
    # playwright.stop()
    # keyboard.press_and_release("ctrl+c")
    # del playwright
    # playwright.dispose()



with sync_playwright() as playwright:
    run(playwright)
