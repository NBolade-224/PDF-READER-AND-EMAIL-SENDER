from tkinter import *
from tkinter import filedialog
from pathlib import Path
import win32com.client as win32
import pytesseract, os, PyPDF3, openpyxl, sys, shutil, time
from pdf2image import convert_from_path
from PIL import ImageTk, Image

MainWindow=Tk()
MainWindow.title("Sending Emails") 
MainWindow.geometry("600x700+600+200") 
MainWindow.configure(bg="RoyalBlue4")

PopplerPath = os.getcwd()+'\\poppler-0.68.0_x86\\poppler-0.68.0\\bin' 
pytesseract.pytesseract.tesseract_cmd =  os.getcwd()+'\\Tesseract-OCR\\tesseract.exe' 

###### SAFETY TO MAKE SURE THAT EACH BUTTON IS PRESSED IN ORDER ########
global Excel_File_Selected_Safety
global PDF_Folder_Select_Safety
Excel_File_Selected_Safety = True
PDF_Folder_Select_Safety = True

def browser():
    ###### LABELS 
    labelforERRORs.configure(text='')
    button_explore.configure(bg='red3')

    ###### GLOBALS 
    global Excel_File_Selected_Safety
    global NameListToEmailListDict

    ###### LISTS. THIS LISTS ARE ORDERED AND WILL BE ZIPPED INTO DICTIONARY
    EmailAddressList = [] 
    NameList = [] 

    ## THIS OPENS THE EXCEL FILE. IF THE FILE IS AN ERROR (FALSE), THEN THE EXCEPTION WILL RETURN AN ERROR LABEL TELLING THE APP USER TO SELECT THE CORRECT FILE
    try:
        filename = filedialog.askopenfilename(parent=MainWindow, initialdir = "C:\\Users\\nickb\\Desktop\\",title = "Select file",filetypes = (("Excel","*xlsx"),))
        ExcelFile = openpyxl.load_workbook(filename)
        ExcelSheet = ExcelFile.active
        filename0 = Path(filename).stem 
        labelfileopened.configure(text="File Opened: "+filename0)
    except:
        labelforERRORs.configure(text='Please select a valid excel file')
        Excel_File_Selected_Safety = True
        return
    
    ## AS STATED EARLIER, THIS CHECKS THAT THE 'B' COLUMN CONTAINS EMAIL ADDRESSES AND RETURNS AN ERROR LABEL IF FALSE
    AddressListForEmails = []
    for cell in ExcelSheet['B']:
        if cell.value != None and '@' in cell.value:
            AddressListForEmails.append(cell.value)
    TotalEmailAdressesPDF = len(AddressListForEmails)
    if TotalEmailAdressesPDF > 0:
        AddressListForEmails = []
        pass
    else:
        labelforERRORs.configure(text='ERROR: No email address found in Col "B"\nPlease select correct Excel File')
        MainWindow.update()
        Excel_File_Selected_Safety = True
        return

    ## AFTER THE ABOVE EXCEPTIONS HAVE BEEN RESOLVED, THE LISTS FROM EARLIER WILL NOW BE POPULATED WITH THE DATA, AS SEEN BELOW
    for cell in ExcelSheet['A']:
        NameList.append(cell.value)
    for cell in ExcelSheet['B']:
        EmailAddressList.append(cell.value)

    ## LISTS ARE NOW ZIPPED TOGETHER TO CREATE A KEY/VALUE SEARCH FORMAT
    NameListToEmailListDict = dict(zip(NameList, EmailAddressList))

    ## PDF WINDOW PROPERTIES + SAFETY
    button_explore.configure(bg='green4')
    Excel_File_Selected_Safety = False
    MainWindow.update()

def CheckMainFolder():
    global Excel_File_Selected_Safety
    global PDF_Folder_Select_Safety
    global PDFs_In_Folder_To_Sort 
    global Total_PDFs_To_Sort
    global PDF_Folder_To_Sort    

    ## SAFETY CHECK
    labelforERRORs.configure(text="")
    button_explore2.configure(bg='red3')
    PDF_Folder_Select_Safety = True
    if Excel_File_Selected_Safety == True:
        labelforERRORs.configure(text="Please select a leavers report first")
        return

    ## THIS OPENS THE EXCEL FILE. IF THE FILE IS AN ERROR (FALSE), THEN THE EXCEPTION WILL RETURN AN ERROR LABEL TELLING THE APP USER TO SELECT THE CORRECT FILE
    PDF_Folder_To_Sort = filedialog.askdirectory(parent=MainWindow, initialdir = "C:\\Users\\nickb\\Desktop\\Test Email Sending",title = "Select folder")
    if PDF_Folder_To_Sort == "":
        labelforERRORs.configure(text="Please select a folder")
        return

    ## THIS LISTS ALL FILES IN THE ABOVE FOLDER (WHICH IS THE ASSIGNED FOLDER FOR SORTING). THE REASON FOR THIS, IS DUE TO TKINTER NOT ALLOWING TO VIEW FILES IN DIALOG/FOLDER SELECTION
    os.listdir(PDF_Folder_To_Sort)
    PDFs_In_Folder_To_Sort = []
    Path(PDF_Folder_To_Sort+'\\PDFs Sent').mkdir(parents=True, exist_ok=True)
    Path(PDF_Folder_To_Sort+'\\PDFs Ready To Send').mkdir(parents=True, exist_ok=True)

    for files in os.listdir(PDF_Folder_To_Sort):
        if files.endswith('.pdf') or files.endswith('.PDF'):
            PDFs_In_Folder_To_Sort.append(files)
    
    Total_PDFs_To_Sort = len(PDFs_In_Folder_To_Sort)

    window1 = Toplevel(MainWindow)
    window1.title('All PDF files in folder, please select another folder if wrong')
    window1.geometry("900x700+500+200") 
    window1.configure(bg="RoyalBlue4")
    window1.attributes('-topmost',1)

    ListofPDFfilesinFolder = Listbox(window1, bg="royalblue1",width=80, height=31, selectmode='single', font=('Times', 14))
    ListofPDFfilesinFolder.pack()
    ListofPDFfilesinFolder.insert(END,'All PDF files in selected folder, please choose another folder if wrong')
    ListofPDFfilesinFolder.insert(END,'')
    ListofPDFfilesinFolder.insert(END,'Total files in folder: '+str(Total_PDFs_To_Sort))
    ListofPDFfilesinFolder.insert(END,'')

    for x in PDFs_In_Folder_To_Sort:
        ListofPDFfilesinFolder.insert(END,x)

    labetotalPDFtosend.configure(text="PDFs to be processed: "+str(Total_PDFs_To_Sort))

    PDF_Folder_Select_Safety = False
    
    button_explore2.configure(bg='green4')
    MainWindow.update()

def Sort():
    global Excel_File_Selected_Safety
    global PDF_Folder_Select_Safety

    ## SAFETY CHECKS
    if Excel_File_Selected_Safety == True:
        labelforERRORs.configure(text="Please select a leavers report before sorting")
        MainWindow.update()
        return
    
    if PDF_Folder_Select_Safety == True:
        labelforERRORs.configure(text="Please check the PDF folder before sorting")
        MainWindow.update()
        return

    labelforERRORs.configure(text="")
    labetotalPDFtosend.configure(text="")
    FileofErrors = 0

    for currentPDFiter,PDFs in enumerate(PDFs_In_Folder_To_Sort):
        currentPDFiteration = currentPDFiter + 1
        ## RESETTING ALL VARIABLES
        MatchedName = ""

        page = convert_from_path(PDF_Folder_To_Sort+"\\"+PDFs, 450,poppler_path = PopplerPath) 
        page[0].save(PDFs[:-4]+'.jpg', 'JPEG')
        x = Image.open(PDFs[:-4]+'.jpg')
        pageContent = pytesseract.image_to_string(x)

        for content in pageContent.split():
            if content in NameListToEmailListDict:
                MatchedName = content
                break

        if MatchedName == "":
            FileofErrors += 1
            x.close()
            os.remove(PDFs[:-4]+'.jpg')
            labetotalPDFtosend.configure(text='File being Processed: '+str(currentPDFiteration)+'/'+str(Total_PDFs_To_Sort)+"\n Files Unable to Rename: "+str(FileofErrors)+'/'+str(Total_PDFs_To_Sort))
            os.rename(PDF_Folder_To_Sort+"\\"+PDFs,PDF_Folder_To_Sort+"\\"+PDFs[:-4]+" NO MATCH FOUND"+PDFs[-4:])
            MainWindow.update()
            continue
        
        x.close()

        os.remove(PDFs[:-4]+'.jpg') #
        PDFFile = PyPDF3.PdfFileReader(PDF_Folder_To_Sort+'/'+PDFs)
        NumberOfPages = PDFFile.numPages
        Output_PDFFile = PyPDF3.PdfFileWriter()

        for i in range(NumberOfPages):
            Output_PDFFile.addPage(PDFFile.getPage(i))
        Output_PDFFile.encrypt(MatchedName) 
        Output_PDFFile.write(open(PDF_Folder_To_Sort+'\\PDFs Ready To Send\\'+MatchedName+" .pdf", 'wb'))

        labetotalPDFtosend.configure(text='File being Processed: '+str(currentPDFiteration)+'/'+str(Total_PDFs_To_Sort)+"\n Files Unable to Rename: "+str(FileofErrors)+'/'+str(Total_PDFs_To_Sort))
        os.remove(PDF_Folder_To_Sort+"\\"+PDFs)
        MainWindow.update()
    
    
    ## PDF WINDOW PROPERTIES
    labetotalPDFtosend.configure(text='Total files renamed: '+str(currentPDFiteration)+'/'+str(Total_PDFs_To_Sort)+"\n Files Unable to Rename: "+str(FileofErrors)+'/'+str(Total_PDFs_To_Sort))
    button_sort.configure(bg='green4')
    MainWindow.update()

def Send():
    global Excel_File_Selected_Safety
    global PDF_Folder_Select_Safety

    # SAFETY CHECK
    if Excel_File_Selected_Safety == True or PDF_Folder_Select_Safety == True:
        labelforERRORs.configure(text="Please select a leavers report & check folder before sending")
        MainWindow.update()
        return
    
    PDFstoSend = os.listdir(PDF_Folder_To_Sort+'\\PDFs Ready To Send')
    ListofPDFstoEmailAddress = []

    ## LOOP FOR SENDING ALL FILES IN THE FOLDER LISTED ABOVE (WHICH IN TURN, IS AUTOSET/CREATED)
    for files4 in PDFstoSend:
        if files4.endswith('.pdf') or files4.endswith('.PDF'):
            EmailAddress = ""
            NameOnthePDFFile = ''
            NameOnthePDFFile = Path(files4).stem 
            EmailAddress = NameListToEmailListDict.get(NameOnthePDFFile[:-1])
            if EmailAddress == None:
                labelforERRORs.configure(text=str(NameOnthePDFFile)+" Match not found on Excel File")
                MainWindow.update()
                return
            ListofPDFstoEmailAddress.append(NameOnthePDFFile+'    ---->    '+EmailAddress)
            NameOnthePDFFile = ''
            EmailAddress = "" 

    ## THIS GENERATED A NEW WINDOW TO SHOW WHERE EACH PDF IS GOING TO BE SENT, BEFORE SENDING
    window2 = Toplevel(MainWindow)
    window2.title('List of PDFs and Address for them to be sent to')
    window2.geometry("900x700+500+200") 
    window2.configure(bg="RoyalBlue4")
    window2.attributes('-topmost',1)
    ListofPDFtoEmail = Listbox(window2, bg="royalblue1",width=80, height=25, selectmode='single', font=('Times', 14))
    ListofPDFtoEmail.pack()
    TotalToBeSent = len(PDFstoSend)
    ListofPDFtoEmail.insert(END,'Total PDFs to be sent '+str(len(PDFstoSend)))
    ListofPDFtoEmail.insert(END,'')
    for x in ListofPDFstoEmailAddress:
        ListofPDFtoEmail.insert(END,x)
    
    ## THIS STARTS SENDING OUT ALL THE PDFS
    def SendSecond():
        global Excel_File_Selected_Safety
        PDFsentCOunter = 0
        for PDFstos in PDFstoSend:
            if PDFstos.endswith('.pdf') or PDFstos.endswith('.PDF'):
                MatchedName = ""
                EmailAddress = ""
                file = Path(PDF_Folder_To_Sort+'\\PDFs Ready To Send\\'+PDFstos).stem      
                MatchedName = file[:-1]
                EmailAddress = NameListToEmailListDict.get(MatchedName)
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = EmailAddress
                mail.Subject = 'PDF'
                mail.Body = 'Please find your PDF enclosed in this email.'
                mail.Attachments.Add(PDF_Folder_To_Sort+'\\PDFs Ready To Send\\'+PDFstos)
                #mail.CC = ""
                mail.Send()
                PDFsentCOunter += 1
                LabelSendSecondForSent.configure(text = "Amount of files sent: %d/%d" % (PDFsentCOunter,TotalToBeSent))
                MatchedName = ""
                EmailAddress = ""
                shutil.move(PDF_Folder_To_Sort+'\\PDFs Ready To Send\\'+PDFstos,PDF_Folder_To_Sort+'\\PDFs Sent\\'+PDFstos)
                MainWindow.update()
                time.sleep(2)
        LabelSendSecondForSent.configure(text = "All emails sent %d/%d\n\nPlease remove all files in 'PDFs To Send', ready for next process" % (PDFsentCOunter,TotalToBeSent))
        Excel_File_Selected_Safety = True
        MainWindow.update()

    SendPDFMain2135 = Button(window2,text = "Send",bg = 'blue',width = 30,height = 2,command = SendSecond, font=('Times', 15, 'bold'), fg = "yellow2")
    LabelSendSecondForSent = Label(window2,text = "",width = 75,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
    LabelSendSecondForSent.pack()
    SendPDFMain2135.pack()

## PDF WINDOW PROPERTIES
label_file_explorer = Label(MainWindow,text = "Original Factory Shop Email Sender - By Nick",width = 100, height = 4,fg = "white",bg = 'RoyalBlue4', font=('Times', 13)) #RoyalBlue4
button_explore = Button(MainWindow,text = "Select Chris21 Leavers Report",bg = 'red3',width = 30, height = 2,command = browser, font=('Times', 15, 'bold'), fg = "yellow2")
button_explore2 = Button(MainWindow,text = "Select To List All Files In Folder",bg = 'red3',width = 30, height = 2,command = CheckMainFolder, font=('Times', 15, 'bold'), fg = "yellow2")
button_exit = Button(MainWindow,text = "Exit",bg = 'snow4',width = 30,height = 2,command = sys.exit, font=('Times', 15, 'bold'), fg = "black")
button_sort = Button(MainWindow,text = "Sort",bg = 'blue',width = 30,height = 2,command = Sort, font=('Times', 15, 'bold'), fg = "yellow2")
button_send = Button(MainWindow,text = "Send",bg = 'blue',width = 30,height = 2,command = Send, font=('Times', 15, 'bold'), fg = "yellow2")
labelfileopened = Label(MainWindow,text = "",width = 75, height = 2,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
labetotalPDFtosend = Label(MainWindow,text = "",width = 50, height = 2,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
labelforERRORs = Label(MainWindow,text = "",width = 75, height = 2,fg = "red",bg = 'RoyalBlue4', font=('Times', 16))
labetotalPDFtosend = Label(MainWindow,text = "",width = 75,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
LabelSpace1PDF = Label(MainWindow,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')
LabelSpace2PDF = Label(MainWindow,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')
LabelSpace3PDF = Label(MainWindow,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')
LabelSpace4PDF = Label(MainWindow,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')

TitleImage = os.getcwd()+'\\1519797862804.jpg'
img = Image.open(TitleImage)
img = img.resize((500, 100), Image.LANCZOS)
img = ImageTk.PhotoImage(img)
labelimage = Label(MainWindow, image = img,width = 500, height = 100)

button_send.configure(bg='blue')

labelimage.pack()
label_file_explorer.pack()
button_explore.pack()
LabelSpace1PDF.pack()
button_explore2.pack()
LabelSpace2PDF.pack()
button_sort.pack()
LabelSpace3PDF.pack()
button_exit.pack()
LabelSpace4PDF.pack()
button_send.pack()
labelforERRORs.pack()
labelfileopened.pack()
labetotalPDFtosend.pack()
labetotalPDFtosend.pack()

MainWindow.mainloop()
