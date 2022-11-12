
import tkinter as tk
import tkinter.messagebox
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
        self.export_to_xlsx_btn['command'] = self.export_last_run
        self.export_to_xlsx_btn.pack(side=tk.BOTTOM ,expand=True)
        self.export_to_xlsx_btn['state'] = 'disable'

        #----------------Other Gui elements -----------------

        #Scrolled text widget for logging
        self.log_area = ScrolledText(self.main_win, width=500)
        self.log_area.insert(tk.INSERT,"Please select screenshots to parse..")
        self.log_area.pack(fill=tk.BOTH , expand=True)
    
    
    # Uses tesserocr to print info
    def get_data_from_files(self):
        #Clean last run cache
        self.last_run = []

        #Ask for screenshot files to parse
        img_files = tkinter.filedialog.askopenfilenames(initialdir='./')

        #Send message that no file was selected
        if(len(img_files) == 0):
            tkinter.messagebox.showerror('Selection error', 'Error: Files selected')
            return None
        

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
                #Store these in variables
                self.log_area.insert(tk.INSERT,'\n{}\n'.format(img))
                date_str = ''
                amnt = -1
                ref_str = ''

                #Flags to tell me if the key info has been found
                fnd_amt = False
                fnd_ref = False
                fnd_dat = False


                for line in raw_list:
                    #If amount has not been parsed just yet then look for it first
                    if(not fnd_amt):

                        if( (line.find('Amount Due PHP') != -1) or \
                            (line.find('Amount Paid PHP') != -1) or \
                            (line.find('Total PHP') != -1)  ):

                            # Parse string containing figure : 
                            # - remove trailing spaces and commas 
                            # - convert to float
                            amnt = float( line[line.find('PHP') + len('PHP'):].strip().replace(',',''))

                            #Print for debugging
                            self.log_area.insert(tk.INSERT, 'Amount : PHP {}'.format(amnt) + '\n')
                            
                            #Assert flag
                            fnd_amt = True

                    elif((line.find('Ref. No.') != -1) and ~fnd_ref):
                        # Parse string containing ref. no. : 
                        # - remove trailing spaces and commas                      
                        ref_str = line[line.find('Ref. No.') + len('Ref. No.'): ].strip().replace(' ','')
                        self.log_area.insert(tk.INSERT, 'Ref. No. : {} \n'.format(ref_str))
                        fnd_ref = True

                    else:
                        for month in self.months:
                            if((line.find(month) != -1) and ~fnd_dat ):
                                date_str = line
                                # Parse string containing date : 
                                # - remove trailing spaces and commas                      
                                self.log_area.insert(tk.INSERT, 'Date : {}\n'.format(date_str))
                                fnd_dat = True
                    
                        
                    if(fnd_dat & fnd_amt and fnd_ref):
                        #Store in last run cache
                        self.last_run.append([date_str, amnt,ref_str])

                        #Clean up                    
                        fnd_dat = False
                        fnd_amt = False
                        fnd_ref = False
                        
                        #Reset flags
                        amnt = -1
                        ref_str = ''
                        date_str = '' 

                #Change state of export to xlsx button after first run
                if(self.export_to_xlsx_btn['state'] != 'normal'):
                    self.export_to_xlsx_btn['state'] = 'enable'
                      

                self.log_area.insert(tk.INSERT,'\n--------------------------\n')
                
            
        self.log_area.insert(tk.INSERT,'Done! '.format(len(img_files)))
        print('Done')
                        
    def export_last_run(self):
        output_filename =  tkinter.filedialog.asksaveasfilename( initialdir='./', defaultextension=".xlsx")
        
        #Send message that no file was selected
        if(len(output_filename) == 0):
            tkinter.messagebox.showerror('Selection error', 'Error: No filename input')
            return None
            
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(output_filename)
        worksheet = workbook.add_worksheet()

        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0

        # Iterate over the data and write it out row by row.
        for date, amnt, ref in (self.last_run):
            worksheet.write(row, col,     date)
            worksheet.write(row, col + 1, amnt)
            worksheet.write(row, col + 2, ref)
            row += 1

        workbook.close()


if __name__ == "__main__":
    app = Gcash_parser()
    app.mainloop()
