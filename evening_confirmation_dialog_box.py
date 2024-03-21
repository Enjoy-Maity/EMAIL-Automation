import re
import os
import sys
import tkinter
from tkinter import *
import tkinter.ttk as ttk
from tkinter import messagebox
from PIL import ImageTk, Image


class Confirmation_dialog_box(tkinter.Tk):
    def __init__(self, parent, **kwargs):
        self.result_dictionary = {}
        self.var = 0
        # if parent is None:
        #     parent = Tk()
        self.parent = parent
        # if self.parent.main_win.state() == "normal":
        #     self.parent.main_win.withdraw()
        self.top = Toplevel(self.parent.main_win)
        # if self.parent.main_win.state() == 'normal':
            # self.parent.main_win.withdraw()
        if 'iconified' not in self.parent.main_win.wm_state():
            self.parent.main_win.iconify()
        self.total_planned_crs = None
        self.total_night_executors = None
        self.total_resources_on_leaves = None
        self.total_resources_on_comp_off = None
        if len(kwargs) > 0:
            if "total_planned_crs" in kwargs:
                self.total_planned_crs = kwargs["total_planned_crs"]

            if "total_night_executors" in kwargs:
                self.total_night_executors = kwargs["total_night_executors"]

            if "total_resources_on_leaves" in kwargs:
                self.total_resources_on_leaves = kwargs["total_resources_on_leaves"]

            if "total_resources_on_comp_off" in kwargs:
                self.total_resources_on_comp_off = kwargs["total_resources_on_comp_off"]

        self.top.title("    Confirmation Dialog Box")
        self.top.geometry("600x550")
        self.top.maxsize(600, 550)
        self.top.minsize(600, 550)
        self.top.iconbitmap(".\\images\\ericsson-blue-icon-logo.ico")
        self.top.bind("<Escape>", self.close_window)
        self.top.bind("<Alt-F4>", self.close_window)
        self.top.bind("<Return>", self.submit)


        self.accidental_closed_flag = 0

        self.style = ttk.Style()
        self.style.theme_use('vista')
        self.style.theme_settings("vista", {
            "TButton": {
                "configure": {
                    "padding": 2,
                    "font": "Ericsson_Hilda 12 bold"
                }
            },
            "TMenubutton": {
                "configure": {
                    "font": "Ericsson_Hilda 12",
                    'justify': 'center',
                    'width': 20
                }
            }
        })

        self.canvas = Canvas(self.top, width=600, height=550, highlightthickness=0, relief='ridge', bd=0)
        self.canvas.grid(column = 0, row = 0, sticky=NW)

        self.top.grab_set()
        self.top.focus_force()

        self.top.resizable(False,False)
        self.top.protocol("WM_DELETE_WINDOW", self.close_window)

        self.font_style = ("Ericsson Hilda", 17, "bold")
        self.font_style_1 = ("Ericsson Hilda", 12)

        self.image = ImageTk.PhotoImage(Image.open(".\\images\\MPBN PLANNING TASK_3_4.png"))
        self.canvas.create_image(0, 0, anchor = "nw", image = self.image)

        # self.canvas.create_text(130, 30, text="Total Planned CRs :-", font=self.font_style, fill="white")
        self.canvas.create_text(130, 70, text="Total Planned CRs :-", font=self.font_style, fill="white")
        self.planned_crs = Entry(self.top, width=25, font=self.font_style_1, bd=0, relief="ridge", justify="center")
        if self.total_planned_crs is not None:
            self.planned_crs.insert(0, self.total_planned_crs)
        # self.canvas.create_window(450, 27, window=self.planned_crs)
        self.canvas.create_window(450, 67, window=self.planned_crs)

        # self.canvas.create_text(122, 70, text="Total Picked CRs :-", font=self.font_style, fill="white")
        self.canvas.create_text(122, 30, text="Total Picked CRs :-", font=self.font_style, fill="white")
        self.picked_crs = Entry(self.top, width=25, font=self.font_style_1, bd=0, relief="ridge", justify="center")
        if self.total_planned_crs is not None:
            self.picked_crs.insert(0, self.total_planned_crs)
        # self.canvas.create_window(450, 67, window=self.picked_crs)
        self.canvas.create_window(450, 27, window=self.picked_crs)

        self.canvas.create_text(135, 110, text="Total Day Planners :-", font=self.font_style, fill="white")
        self.day_planners_var = IntVar(self.top)
        values = [i for i in range(1, 16)]
        self.day_planners_file_check()
        if (self.var is not None) or (isinstance(self.var, int)):
            if self.var != 0:
                self.day_planners_var.set(self.var)
                # print("self.var",self.var)
            else:
                self.day_planners_var.set(3)

        if self.var is None:
            self.day_planners_var.set(3)

        # print(self.day_planners_var.get())
        self.day_planners = ttk.OptionMenu(self.top, self.day_planners_var, *values, style="TMenubutton")
        self.canvas.create_window(450, 108, window=self.day_planners)

        # self.canvas.create_text(140, 150, text="Total Team Resources", font=self.font_style, fill="white")
        self.canvas.create_text(150, 150, text="Total Night Resources :-", font=self.font_style, fill="white")
        values_for_night_executors = [i for i in range(0, 26)]
        self.total_night_executors_var = IntVar(self.top)
        # Entering the total night executors number in the option menu variable
        if self.total_night_executors is not None:
            self.total_night_executors_var.set(self.total_night_executors)
        self.total_night_executors_optionmenu = ttk.OptionMenu(self.top, self.total_night_executors_var, *values_for_night_executors, style="TMenubutton")
        self.canvas.create_window(450, 148, window=self.total_night_executors_optionmenu)

        # self.canvas.create_text(172, 190, text="Total Resources on Leaves :-", font=self.font_style, fill="white")
        self.canvas.create_text(172, 190, text="Total Resources on Leaves :-", font=self.font_style, fill="white")
        values_for_resources_on_leaves = [0]
        values_for_resources_on_leaves.extend([i for i in range(0, 26)])
        self.total_resources_on_leaves_var = IntVar(self.top)
        # Entering the total resources on leaves number in the option menu variable
        if self.total_resources_on_leaves is not None:
            self.total_resources_on_leaves_var.set(self.total_resources_on_leaves)
        self.total_resources_on_leaves_optionmenu = ttk.OptionMenu(self.top, self.total_resources_on_leaves_var, *values_for_resources_on_leaves, style="TMenubutton")
        # self.canvas.create_window(450, 188, window=self.total_resources_on_leaves_optionmenu)
        self.canvas.create_window(450, 188, window=self.total_resources_on_leaves_optionmenu)

        # self.canvas.create_text(172, 310, text="Total Resources on Comp-Off :-", font=self.font_style, fill="white")
        self.canvas.create_text(172, 230, text="Total Resources on Comp-Off :-", font=self.font_style, fill="white")
        values_for_resources_on_comp_off = [0]
        values_for_resources_on_comp_off.extend([i for i in range(0, 26)])
        self.total_resources_on_comp_off_var = IntVar(self.top)
        # Entering the total resources on leaves number in the option menu variable
        if self.total_resources_on_comp_off is not None:
            self.total_resources_on_comp_off_var.set(self.total_resources_on_leaves)
        self.total_resources_on_comp_off_optionmenu = ttk.OptionMenu(self.top, self.total_resources_on_comp_off_var, *values_for_resources_on_comp_off, style="TMenubutton")
        # self.canvas.create_window(450, 308, window=self.total_resources_on_comp_off_optionmenu)
        self.canvas.create_window(450, 228, window=self.total_resources_on_comp_off_optionmenu)

        # self.canvas.create_text(300, 240, text="Resources who are on Leave :-", font=self.font_style, fill="white")
        self.canvas.create_text(300, 290, text="Resources who are on Leave :-", font=self.font_style, fill="white")
        self.resources_on_leaves_entry = ttk.Entry(self.top, width=50, font=self.font_style_1, justify="center")
        # self.canvas.create_window(290, 270, window=self.resources_on_leaves_entry)
        self.canvas.create_window(290, 320, window=self.resources_on_leaves_entry)

        # self.canvas.create_text(300, 360, text="Resources who are on Comp-off :-", font=self.font_style, fill="white")
        self.canvas.create_text(300, 370, text="Resources who are on Comp-off :-", font=self.font_style, fill="white")
        self.resources_on_comp_off_entry = ttk.Entry(self.top, width=50, font=self.font_style_1, justify="center")
        # self.canvas.create_window(290, 390, window=self.resources_on_comp_off_entry)
        self.canvas.create_window(290, 400, window=self.resources_on_comp_off_entry)

        # self.top.destroy()
        # parent.deiconify(
        #
        # )

        self.submit_button = ttk.Button(self.top, text="Submit", style="TButton", command=lambda: self.submit(""))
        self.canvas.create_window(450, 450, window=self.submit_button)

        if self.top.state() != "normal":
            print("returning from here")
            # return self.get_details()
        self.top.mainloop()

    def close_window(self, *args):
        self.accidental_closed_flag = 1
        # self.parent.deiconify()
        self.top.withdraw()
        # self.parent.main_win_flag = 0
        # self.parent.main_win.deiconify()
        # self.top.destroy()
        self.top.quit()
        # sys.exit(0)

    def submit(self, _):
        day_planners_count = self.day_planners_var.get()

        planned_crs = self.planned_crs.get()
        picked_crs = self.picked_crs.get()
        total_night_executors = self.total_night_executors_var.get()
        resources_who_are_on_leave = self.resources_on_leaves_entry.get()
        resources_who_are_on_comp_off = self.resources_on_comp_off_entry.get()
        number_of_resources_who_are_on_leave = self.total_resources_on_leaves_var.get()
        number_of_resources_who_are_on_comp_off = self.total_resources_on_comp_off_var.get()

        compiled_pattern = re.compile(r"[\d!@#%^&*)(\]\[\-/+\"\'=?><|}{~$;:\\]")
        compiled_pattern_one = re.compile(r"[!@#%^&*)(\]\[\-/+\"\'=?><|}{~$;:\\]")
        compiled_pattern_two = re.compile(r"[a-zA-Z!@#%^&*)(\]\[\-/+\"\'=?><|}{~$;:\\]")

        # if re.search(compiled_pattern_two,day_planners_count) is not None:
        #     messagebox.showerror("Error", "Day Planners should not contain any special characters or alphabets")
        #     self.day_planners_var.set("")

        if re.search(compiled_pattern_two, planned_crs) is not None:
            messagebox.showerror("Error", "Planned CRs should not contain any special characters or alphabets")
            self.planned_crs.delete(0, END)

        if re.search(compiled_pattern_two, picked_crs) is not None:
            messagebox.showerror("Error", "Picked CRs should not contain any special characters or alphabets")
            self.picked_crs.delete(0, END)

        if re.search(compiled_pattern, resources_who_are_on_comp_off) is not None:
            messagebox.showerror("Error", "Resources who are on comp-off should not contain any special characters or numbers")
            self.resources_on_comp_off_entry.delete(0, END)

        if re.search(compiled_pattern, resources_who_are_on_leave) is not None:
            messagebox.showerror("Error", "Resources who are on leave should not contain any special characters or numbers")
            self.resources_on_comp_off_entry.delete(0, END)

        if (number_of_resources_who_are_on_comp_off == 0) and (len(resources_who_are_on_comp_off) != 0):
            messagebox.showerror("Error", "Number of resources who are on comp-off and entry for resources who are on comp-off mismatch")
            self.resources_on_comp_off_entry.delete(0, END)

        if (number_of_resources_who_are_on_leave == 0) and (len(resources_who_are_on_leave) != 0):
            messagebox.showerror("Error", "Number of resources who are on leave and entry for resources who are on leave mismatch")
            self.resources_on_leaves_entry.delete(0, END)

        if (isinstance(planned_crs, int)) and (int(planned_crs) == 0):
            messagebox.showerror("Error", "Planned CRs cannot be 0")
            self.planned_crs.delete(0, END)

        if (isinstance(picked_crs, int)) and (int(picked_crs) == 0):
            messagebox.showerror("Error", "Picked CRs cannot be 0")
            self.picked_crs.delete(0, END)

        if int(total_night_executors) == 0:
            messagebox.showerror("Error", "Total night executors cannot be 0")

        # print(f"{number_of_resources_who_are_on_comp_off = }")
        list_of_resources_who_are_on_comp_off = resources_who_are_on_comp_off.split(",")
        if len(list_of_resources_who_are_on_comp_off) == 1:
            if str(list_of_resources_who_are_on_comp_off[0]).strip().upper() in ("NA", "N/A", ""):
                list_of_resources_who_are_on_comp_off = []
        # print(f'{len(list_of_resources_who_are_on_comp_off) =}')
        if number_of_resources_who_are_on_comp_off != len(list_of_resources_who_are_on_comp_off):
            messagebox.showerror("Error", "Number of resources who are on comp-off and entry for resources who are on comp-off mismatch")
            self.resources_on_comp_off_entry.delete(0, END)

        list_of_split_resources_who_are_on_leave = resources_who_are_on_leave.split(",")
        if len(list_of_split_resources_who_are_on_leave) == 1:
            if str(list_of_split_resources_who_are_on_leave[0]).strip().upper() in ("NA", "N/A", ""):
                list_of_split_resources_who_are_on_leave = []

        if number_of_resources_who_are_on_leave != len(list_of_split_resources_who_are_on_leave):
            messagebox.showerror("Error", "Number of resources who are on leave and entry for resources who are on leave mismatch")
            self.resources_on_leaves_entry.delete(0, END)

        else:
            if self.var != day_planners_count:
                self.day_planners_file_update()
            self.result_dictionary = {
                "planned_crs": planned_crs,
                "picked_crs": picked_crs,
                "total_night_executors": total_night_executors,
                "resources_who_are_on_leave": resources_who_are_on_leave,
                "resources_who_are_on_comp_off": resources_who_are_on_comp_off,
                "number_of_resources_who_are_on_leave": number_of_resources_who_are_on_leave,
                "number_of_resources_who_are_on_comp_off": number_of_resources_who_are_on_comp_off,
                "day_planners": day_planners_count
            }
            self.top.withdraw()
            # self.parent.main_win_flag = 0
            # self.parent.main_win.deiconify()
            # self.top.destroy()
            self.top.quit()

        # sys.exit(0)

    def day_planners_file_check(self):
        username = os.popen(cmd='cmd.exe /C "echo %USERNAME%"').read().strip()
        self.var = 0
        if os.path.exists(f"C:\\Users\\{username}\\AppData\\Local\\MPBN_Planning_Task\\day_planners.txt"):
            with open(f"C:\\Users\\{username}\\AppData\\Local\\MPBN_Planning_Task\\day_planners.txt", 'r') as f:
                self.var = int(f.readline())
                f.close()

            del f

        if isinstance(self.var,int):
            if self.var != 0:
                print(self.var)
                self.day_planners_var.set(self.var)
                print(self.day_planners_var.get())
            else:
                self.day_planners_var.set(3)
        else:
            self.day_planners_var.set(3)

    def day_planners_file_update(self):
        username = os.popen(cmd='cmd.exe /C "echo %USERNAME%"').read().strip()
        if not os.path.exists(f"C:\\Users\\{username}\\AppData\\Local\\MPBN_Planning_Task\\"):
            os.mkdir(f"C:\\Users\\{username}\\AppData\\Local\\MPBN_Planning_Task\\")

        with open(f"C:\\Users\\{username}\\AppData\\Local\\MPBN_Planning_Task\\day_planners.txt", 'w') as f:
            # print(f"Writing {str(int(self.day_planners_var.get()))} in {f'C:/Users/{username}/AppData/Local/MPBN_Planning_Task/day_planners.txt'}")
            f.write(str(int(self.day_planners_var.get())))
            f.close()

        del f

    def get_details(self):
        if self.accidental_closed_flag == 0:
            return self.result_dictionary
        else:
            return "Unsuccessful"

# app = Tk()
# Confirmation_dialog_box(app)
# app.mainloop()
