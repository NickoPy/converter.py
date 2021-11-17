from colorama import init
from colorama import Fore
import comtypes.client
import time
import datetime
import os
import win32com.client

continua = True

# enter loop
while continua:

    init()
    datetime_object = datetime.datetime.now()
    wdFormatPDF = 17

    # asking what conversion to do
    conversion_input = input(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + Fore.LIGHTBLUE_EX +
                             "] Enter " + Fore.RED + "docx to pdf " + Fore.LIGHTBLUE_EX +
                             "or " + Fore.RED + "pdf to docx " + Fore.LIGHTBLUE_EX + "or " + Fore.RED +
                             "pptx to pdf" + Fore.LIGHTBLUE_EX + ": \n")

    # running script for pdf output
    if conversion_input == "docx to pdf":
        # asking file directory for input and output
        input_file = input(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + Fore.LIGHTBLUE_EX +
                           "] Enter the directory of the file to convert: \n")
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] The chosen directory is " + Fore.RED +
              input_file + Fore.LIGHTBLUE_EX + ". Loading...")
        time.sleep(0.5)
        output_file = input(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + Fore.LIGHTBLUE_EX +
                            "] Enter the directory where to put the converted file: \n")
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] The chosen directory is " + Fore.RED +
              output_file + Fore.LIGHTBLUE_EX + ". Loading...")
        time.sleep(0.5)

        # create COM object
        word = comtypes.client.CreateObject('Word.Application')
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Creating COM object")
        # key point 1: make word visible before open a new document
        word.Visible = True
        # key point 2: wait for the COM Server to prepare well.
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 3s...")
        time.sleep(1)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 2s...")
        time.sleep(1)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 1s...")
        time.sleep(1)

        # convert docx file 1 to pdf file 1
        # open docx file 1
        doc = word.Documents.Open(input_file)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Opening " + Fore.RED + input_file)
        # conversion
        doc.SaveAs(output_file, FileFormat=wdFormatPDF)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Converting docx to pdf...")
        time.sleep(0.5)
        # close docx file 1
        doc.Close()
        word.Visible = False
        # close Word Application
        word.Quit()
        print(Fore.GREEN + "[" + str(datetime_object) + "] File converted successfully :)")
        time.sleep(0.5)

    # running script for docx output
    if conversion_input == "pdf to docx":
        # asking file directory for input and output
        pdf_file = input(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + Fore.LIGHTBLUE_EX +
                         "] Enter the directory of the file to convert: \n")
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] The chosen directory is " + Fore.RED +
              pdf_file + Fore.LIGHTBLUE_EX + ". Loading...")
        time.sleep(0.5)
        docx_file = input(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + Fore.LIGHTBLUE_EX +
                          "] Enter the directory where to put the converted file: \n")
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] The chosen directory is " + Fore.RED +
              docx_file + Fore.LIGHTBLUE_EX + ". Loading...")
        time.sleep(0.5)

        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0

        # getting file name and normalizing path
        filename = pdf_file.split('\\')[-1]
        in_file = os.path.abspath(pdf_file)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Getting ready...")
        time.sleep(0.5)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 3s...")
        time.sleep(1)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 2s...")
        time.sleep(1)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 1s...")
        time.sleep(1)

        # converting to docx and saving
        wb = word.Documents.Open(in_file)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Opening " + Fore.RED + pdf_file)
        time.sleep(0.5)
        out_file = os.path.abspath(docx_file + '\\' + filename[0:-4] + ".docx")
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Converting pdf to docx...")
        time.sleep(0.5)

        wb.SaveAs2(out_file, FileFormat=16)
        wb.Close()
        word.Quit()
        print(Fore.GREEN + "[" + str(datetime_object) + "] File converted successfully :)")
        time.sleep(0.5)

    if conversion_input == "pptx to pdf":
        # asking file directory for input and output
        pptx_file = input(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + Fore.LIGHTBLUE_EX +
                          "] Enter the directory of the file to convert: \n")
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] The chosen directory is " + Fore.RED +
              pptx_file + Fore.LIGHTBLUE_EX + ". Loading...")
        time.sleep(0.5)
        pdf_file = input(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + Fore.LIGHTBLUE_EX +
                         "] Enter the directory where to put the converted file: \n")
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] The chosen directory is " + Fore.RED +
              pdf_file + Fore.LIGHTBLUE_EX + ". Loading...")
        time.sleep(0.5)

        powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
        powerpoint.Visible = 1

        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Getting ready...")
        time.sleep(0.5)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 3s...")
        time.sleep(1)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 2s...")
        time.sleep(1)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Sleeping 1s...")
        time.sleep(1)

        if pdf_file[-3:] != 'pdf':
            pdf_file = pdf_file + ".pdf"
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Opening " + Fore.RED + pptx_file)
        time.sleep(0.5)
        print(Fore.LIGHTBLUE_EX + "[" + str(datetime_object) + "] Converting pptx to pdf...")
        time.sleep(0.5)
        deck = powerpoint.Presentations.Open(pptx_file)
        deck.SaveAs(pdf_file, 32)  # formatType = 32 for ppt to pdf
        deck.Close()
        powerpoint.Quit()
        print(Fore.GREEN + "[" + str(datetime_object) + "] File converted successfully :)")
        time.sleep(0.5)

    # restart/exit loop
    yes_continue = input(Fore.LIGHTBLUE_EX + "Do you want to continue? y/n \n")
    continua = yes_continue == "y"
