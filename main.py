
import tkinter as tk
import tkinter.filedialog
import tkinter.scrolledtext
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk
from turtle import left, width
from tesserocr import PyTessBaseAPI
import xlsxwriter

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

        #Make list to temporarily store the last parse run
        self.last_run = []
        
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

        self.log_area.insert(tk.INSERT,'Selected files: {}\n Starting image recognition'.format(len(img_files)))
        
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
                date_str = ''
                amnt = -1
                ref_str = ''
                for line in raw_list:
                    if(line.find('Amount Due PHP') != -1):
                        amnt = 0
                    elif(line.find('Amount Paid PHP') != -1):
                        amnt = 0
                    elif(line.find('Total PHP') != -1):
                        amnt = 0
                    elif(line.find('Ref. No.') != -1):
                        ref_str = line
                    else:
                        for month in self.months:
                            if(line.find(month) != -1):
                                date_str = line
                    
                    #If Reference number exists, Print and store to data
                    if(len(ref_str) > 0):
                        # Parse string containing ref. no. : 
                        # - remove trailing spaces and commas                      
                        ref_str = line[line.find('Ref. No.') + len('Ref. No.'): ].strip().replace(' ','')
                        self.log_area.insert(tk.INSERT, 'Ref. No. ' + ref_str + '\n')
                        
                    
                    #If Amnt is exists, Print and store to data
                    if(amnt > -1):
                        # Parse string containing figure : 
                        # - remove trailing spaces and commas 
                        # - convert to float
                        amnt = float(line[line.find('PHP') + len('PHP'):].strip().replace(',',''))

                        #Print for debugging
                        self.log_area.insert(tk.INSERT, 'Amount : PHP '+ str(amnt) + '\n')
                        
                    #If date number exists, Print and store to data
                    if(len(date_str) > 0):
                        # Parse string containing date : 
                        # - remove trailing spaces and commas                      
                        self.log_area.insert(tk.INSERT, 'Date : ' + date_str + '\n')
                    
                    #Clean up
                    amnt = -1
                    ref_str = ''
                    date_str = ''

                self.log_area.insert(tk.INSERT,'\n--------------------------\n')
                
            
        self.log_area.insert(tk.INSERT,'Done! '.format(len(img_files)))

                    

        print('Done')
                        
                    


if __name__ == "__main__":
    app = Gcash_parser()
    app.mainloop()
