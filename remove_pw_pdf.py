

from PyPDF2 import PdfReader

ENCRYPTED_FILE_PATH = '.\\transaction_history.pdf'

with open(ENCRYPTED_FILE_PATH, mode='rb') as f:        
    reader = PdfReader(f)
    if reader.is_encrypted:
        reader.decrypt('abanto7019')

        #Test to see if we opened it correctly by counting pages
        print('Number of pages {}'.format(len(reader.pages)))

        reconstructed_lines = []

        for page in reader.pages:
            lines = page.extract_text()
            lines_parsed = lines.split('\n')

            #Declare buffer to repair fragmented lines
            
            is_multiline = 0
            temp_str = ''
            for line in lines_parsed:

                
                #If last character on the string is a space assume that it is a fragment
                if(line[-1] == ' ' and not is_multiline):

                    #Assert flag and store upper fragment
                    is_multiline = 1
                    temp_str = line
                elif(is_multiline):

                    #Concat lower fragment to upper fragment of the line
                    temp_str += line

                    #append reconstructed lines
                    reconstructed_lines.append(temp_str)

                    #De-assert flag 
                    is_multiline = 0
                else:
                    #If line is complete then just append to buffer
                    reconstructed_lines.append(line)
            
        for line in reconstructed_lines:
            print(line)
            