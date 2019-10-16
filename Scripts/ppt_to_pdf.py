#for windows only 
#ppt,pptx to pdf files

import os
from time import strftime
import comtypes.client 

# Counts the number of files in the directory that can be converted
def n_files(directory):
    total = 0
    for file in os.listdir(directory):
        if (file.endswith('.pptx') or file.endswith('.ppt')):
            total += 1
    return total

# Creates a new directory within current directory called PDFs
def createFolder(directory):
    if not os.path.exists(directory + '\\PDFs'):
        os.makedirs(directory + '\\PDFs')

def convert_ppt():
    try:
        ppt = comtypes.client.CreateObject("Powerpoint.Application")
        #ppt.Visible = 1
        for file in os.listdir(directory):
            if (file.endswith('.ppt') or file.endswith('.pptx')):
                ending = ""
                if file.endswith('.ppt'):
                    ending = '.ppt'
                if file.endswith('.pptx'):
                    ending = '.pptx'
                new_name = file.replace(ending,r".pdf")
                in_file = os.path.abspath(directory + '\\' + file)
                new_file = os.path.abspath(directory + '\\PDFs' + '\\' + new_name)
                deck = ppt.Presentations.Open(in_file)
                print(new_name)
                deck.SaveAs(new_file, 32)#file format
                deck.Close()
                ppt.Quit()
    except:
        print("exceptions in convert_ppt")


    

if __name__ == "__main__":
    print('\nPlease note that this will overwrite any existing PDF files')
    print('For best results, close Microsoft Powerpoint before proceeding')
    #input('Press enter to continue.')

    directory = os.getcwd()#gets the current working directory of a process. Can also set custom directory like, directory = 'C:\\Users\\b.mushtaq\\Downloads'

    if n_files(directory) == 0:
        print('There are no files to convert')
        exit()
        
    createFolder(directory)
    print('Starting conversion... \n')
    convert_ppt()
    print('\nConversion finished at ' + strftime("%H:%M:%S"))
    input('press enter to continue')
