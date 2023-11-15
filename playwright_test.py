import os
from playwright.sync_api import Playwright, sync_playwright, expect

def exception_raiser(exception,*args):
    if(exception == 'Exception'):
        if(len(args) == 0):
            raise Exception()

def run(playwright:Playwright) ->None:
    browser_type = playwright.chromium
    browser = browser_type.launch(headless = False, slow_mo = 100)
    context = browser.new_context()

    # Opening the page
    page = context.new_page()

    # Go to the ITSM website
    page.goto("https://nextgentm-in.sdt.ericsson.net/arsys/forms/umt-ars-in/SHR%3ALandingConsole/Default+Administrator+View/?cacheid=a9eea6e3")

    # finding the id 'domain' and 'Employee' in the option value
    select_element = page.locator("select#domain")
    select_element.select_option("Employee")
    page.locator("//a[@id='loginBtn']",has_text="Log On").click()

    # Getting the username via subprocess 
    # cmd /c runs the command in cmd and terminate
    # cmd /k runs the command in cms and remain
    # username = subprocess.run((r"cmd /c",r"'echo %username%'"),capture_output=True,shell=False)
    
    # Getting the username via os.system
    # cmd /c runs the command in cmd and terminate
    # cmd /k runs the command in cms and remain
    # username = subprocess.run((r"cmd /c",r"'echo %username%'"),capture_output=True,shell=False)
    username = os.popen(r'cmd /c "echo %username%"').read().upper()
    
    # Entering the username
    page.locator("//input[@id='login']").fill(username)

    try:
        # Entering the passwd
        page.locator("//input[@id='passwd']").fill("heqt$$65")
        
        # clicking the logon button
        page.locator("//a[@id='loginBtn']",has_text='Log On').click()
        
        print(len(page.get_by_text("Try again after some time or contact your help desk").all()))
        if(len(page.get_by_text("Try again after some time or contact your help desk").all()) > 0):
            raise Exception()
   
    except Exception as e:
        from tkinter import messagebox
        import traceback
        messagebox.showerror("",f"{traceback.format_exc()}\n\n{e}")
    

    context.wait_for_event("close")
    browser.wait_for_event("close")
    playwright.dispose()



with sync_playwright() as playwright:
    run(playwright)