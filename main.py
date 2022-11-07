
import tkinter as tk
import tkinter.filedialog
import tkinter.scrolledtext
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk
from turtle import left, width
from tesserocr import PyTessBaseAPI


class Gcash_parser(tk.Tk):
    def __init__(self):
        #This calls allows the class the inherit tk.Tk()
        #which is similar to simply calling root = Tk()
        super().__init__()

        #Set gui
        self.title('Gcash Screenshot Parser')
        self.geometry('900x500')

        #Make list to lookup month abreviations
        self.months = ['Jan', 'Feb', 'Mar' , 'Apr', 'May', 'Jun', 'Jul', \
                        'Aug', 'Sept', 'Oct', 'Nov', 'Dec']

        
        self.main_win = ttk.Frame(self)
        self.main_win.pack()

        #----------------Buttons for interface--------------
        self.btn_pnl = ttk.Frame(self.main_win)
        self.btn_pnl.pack()

        self.get_files_btn = ttk.Button(self.btn_pnl,text='Select Files')
        self.get_files_btn['command'] = self.get_data_from_files
        self.get_files_btn.pack(side=tk.TOP ,expand=True)

        self.export_to_xlsx_btn = ttk.Button(self.btn_pnl, text='Export to xlsx')
        self.export_to_xlsx_btn.pack(side=tk.BOTTOM ,expand=True)

        #----------------Other Gui elements -----------------

        #Scrolled text widget for logging
        self.log_area = ScrolledText(self.main_win, width=500)
        self.log_area.insert(tk.INSERT,"Please select screenshots to parse..")
        self.log_area.pack(fill=tk.BOTH , expand=True)



    
    
    # Uses tesserocr to print info
    def get_data_from_files(self):

        #Ask for screenshot files to parse
        img_files = tkinter.filedialog.askopenfilenames(initialdir='./')

        
        # Api is automatically finalized when used in a with-statement (context manager).
        # otherwise api.End() should be explicitly called when it's no longer needed.
        with PyTessBaseAPI() as api:
            for img in img_files:

                #use api to convert recognized text to string of chars
                #this is based on the original C++ api
                api.SetImageFile(img)
                raw_text = api.GetUTF8Text()
                raw_list = raw_text.split('\n')
                
                #Print only the valuable information to log
                self.log_area.insert(tk.INSERT,'\n{}\n'.format(img))
                log_str = ''
                for line in raw_list:
                    if(line.find('Amount Due PHP') != -1):
                        log_str = line
                    elif(line.find('Amount Paid PHP') != -1):
                        log_str = line
                    elif(line.find('Total PHP') != -1):
                        log_str = line
                    elif(line.find('Ref. No.') != -1):
                        log_str = line
                    else:
                        for month in self.months:
                            if(line.find(month) != -1):
                                log_str = line
                    if(len(log_str) > 0):
                        self.log_area.insert(tk.INSERT, log_str + '\n')
                        log_str = ''

                self.log_area.insert(tk.INSERT,'\n--------------------------\n')

                    

        print('Done')
                        
                    


if __name__ == "__main__":
    app = Gcash_parser()
    app.mainloop()
