###########################################
# Project : Gcash screenshot parser
# Author: Alfred Abanto
#
# Description:
#   Just a small proof of concept 
#   demonstrating how we can use OCR
#   (Optical Character Recognition)
#   to automate obtaining key information
#   from gcash screenshots, intended
#   for small businesses
#
# License: MIT
# Date: 14/11/2022
###########################################



import threading
import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
import tkinter.scrolledtext
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk
from tesserocr import PyTessBaseAPI
import xlsxwriter
from dateutil.parser import parse

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
                        'Aug', 'Sept', 'Sep', 'Oct', 'Nov', 'Dec']


        #Dictionary to convert months (abbreviated or not) to numeric counterparts
        self.mon2num_dict = {
            'Jan' :         1,
            'January' :     1,
            'Feb' :         2,
            'February' :    2,
            'Mar' :         3,
            'March' :       3,
            'Apr' :         4,
            'April' :       4,
            'May':          5,
            'Jun':          6,
            'June':         6,
            'Jul':          7,
            'July':         7,
            'Aug':          8,
            'August':       8,
            'Sep':          9,
            'Sept':         9,
            'September':    9,
            'Oct':          10,
            'October':      10,
            'Nov':          11,
            'November':     11,
            'Dec':          12,
            'December':     12
        }

        #Make list to temporarily store the last parse run
        self.last_run = []
        
        self.main_win = ttk.Frame(self)
        self.main_win.pack()

        #----------------Settup Buttons for interface--------------
        self.btn_pnl = ttk.Frame(self.main_win)
        self.btn_pnl.pack()

        self.get_files_btn = ttk.Button(self.btn_pnl,text='Select Files')
        self.get_files_btn['command'] = self.multhithread_ocr
        self.get_files_btn.grid(row=0, column=0, sticky='W')

        self.export_to_xlsx_btn = ttk.Button(self.btn_pnl, text='Export to xlsx')
        self.export_to_xlsx_btn['command'] = self.export_last_run
        self.export_to_xlsx_btn.grid(row=0, column=1, sticky='W')
        self.export_to_xlsx_btn['state'] = 'disable'

        #----------------Other Gui elements -----------------

        #Scrolled text widget for logging
        self.log_area = ScrolledText(self.main_win, width=500)
        self.log_area.insert(tk.INSERT,"Please select screenshots to parse..")
        self.log_area.pack(fill=tk.BOTH , expand=True)
    
    
    # Uses tesserocr to print info
    def get_data_from_files(self):
        #Clean last run cache and clear log area
        
        self.last_run = []
        self.log_area.delete('1.0', tk.END)
    

        #Ask for screenshot files to parse
        img_files = tkinter.filedialog.askopenfilenames(initialdir='./')

        #Send message that no file was selected
        if(len(img_files) == 0):
            tkinter.messagebox.showerror('Selection error', 'Error: Files selected')
            #Enable function button again
            if(self.get_files_btn['state'] != tk.NORMAL):
                self.get_files_btn['state'] = tk.NORMAL

            return None
        
        
        
        #Start process
        self.log_area.insert(tk.INSERT,'Selected files: {}\n Starting image recognition'.format(len(img_files)))
        
        # Api is automatically finalized when used in a with-statement (context manager).
        # otherwise api.End() should be explicitly called when it's no longer needed.
        with PyTessBaseAPI(path= './data/tessdata') as api:
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
                    if(((line.find('Amount Due PHP') != -1) or \
                        (line.find('Amount Paid PHP') != -1) or \
                        (line.find('Total PHP') != -1) ) \
                        and not fnd_amt):

                        # Parse string containing figure : 
                        # - remove trailing spaces and commas 
                        # - convert to float
                        amnt = float( line[line.find('PHP') + len('PHP'):].strip().replace(',',''))

                        #Print for debugging
                        self.log_area.insert(tk.INSERT, 'Amount : PHP {}'.format(amnt) + '\n')
                        
                        #Assert flag
                        fnd_amt = True

                    elif((line.find('Ref. No.') != -1) and  not fnd_ref):
                        # Parse string containing ref. no. : 
                        # - remove trailing spaces and commas                      
                        ref_str = line[line.find('Ref. No.') + len('Ref. No.'): ].strip().replace(' ','')
                        self.log_area.insert(tk.INSERT, 'Ref. No. : {} \n'.format(ref_str))
                        fnd_ref = True

                    else:
                        for month in self.months:
                            if((line.find(month) != -1)  and not fnd_dat ):
                                date_str = line

                                # Parse string using dateutil
                                try:
                                
                                    #Try to parse everything using dateutil
                                    str_parsed = parse(date_str,  fuzzy=True)

                                    #If succesful use datetime methods to extract date/time
                                    tm = str_parsed.time().strftime("%I:%M %p")
                                    numeric_date = str_parsed.date()
                                except:
                                    print("Exception:  {}".format(date_str))

                                    #Exception thrown check for errors in OCR

                                    #check if day and year have been concatenated
                                    date_list = date_str.split(' ')
                                    if( (len(date_list[1]) > 2) and (len(date_list) < 4)):
                                        #Assume Error date format will be: %mm %dd%yyyy,%HH%MM
                                        print(date_list)

                                        date_list = date_str.split(',')
                                        numeric_date = date_list[0]
                                        tm = date_list[1]
                                        print('Error date: \'{}\' time: \'{}\''.format(numeric_date,tm))
                                    else:
                                        continue

                                #Uncomment to print values for debugging
                                # print('Numeric date: '+ numeric_date)

                                self.log_area.insert(tk.INSERT, 'Date : {}\n'.format(numeric_date))
                                fnd_dat = True
                     

                    if(fnd_dat and fnd_amt and fnd_ref):
                        #After all key information is found store in cache for export

                        #Get file name only
                        img_filename = img.split('/')
                        img_filename = img_filename[-1]

                        #Store in last run cache
                        self.last_run.append([ img_filename, numeric_date, tm, ref_str ,amnt])

                        #Clean up                    
                        fnd_dat = False
                        fnd_amt = False
                        fnd_ref = False
                        
                        #Reset flags
                        amnt = -1
                        ref_str = ''
                        date_str = '' 

                
                self.log_area.insert(tk.INSERT,'\n--------------------------\n')
                self.log_area.see('end')
                
            
        self.log_area.insert(tk.INSERT,'Done! '.format(len(img_files)))
        

        #Change state of export to xlsx button after first run
        if(self.export_to_xlsx_btn['state'] != tk.NORMAL):
            self.export_to_xlsx_btn['state'] = tk.NORMAL

        #Enable function button again
        if(self.get_files_btn['state'] != tk.NORMAL):
            self.get_files_btn['state'] = tk.NORMAL

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

        #Add headers
        worksheet.write(row, col, 'File name') 
        worksheet.write(row, col + 1, 'Date')
        worksheet.write(row, col + 2, 'Time')
        worksheet.write(row, col + 3, 'Ref. No.')
        worksheet.write(row, col + 4, 'Amount')

        #Increment row index
        row += 1
        # Iterate over the data and write it out row by row.
        for  filename, date, time, ref, amnt in (self.last_run):
            worksheet.write(row, col,     filename)
            worksheet.write(row, col + 1, date)
            worksheet.write(row, col + 2, time)
            worksheet.write(row, col + 3, ref)
            worksheet.write(row, col + 4, amnt)
            row += 1

        workbook.close()

    #Use multithreading to prevent process from blocking the main program window
    def multhithread_ocr(self):
        #Create new thread for printing
        t1 = threading.Thread(target=self.get_data_from_files )
        t1.start()

        #Disable button to prevent users from running while we're parsing
        self.get_files_btn['state'] = tk.DISABLED


if __name__ == "__main__":
    app = Gcash_parser()
    app.mainloop()
