import os
import sys
import pandas
# Import System libraries
import glob
import random
import re
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

sys.coinit_flags = 0 # comtypes.COINIT_MULTITHREADED

# USE COMTYPES OR WIN32COM
#import comtypes
#from comtypes.client import CreateObject

# USE COMTYPES OR WIN32COM
import win32com
from win32com.client import Dispatch
def macrosExcel(filePath):
    global fname
    fname =filePath
    scripts_dir = r"E:\Python\BM_LR_Template-master\BM_LR_Template-master\{}".format(filePath)
    print("Hello")
    strcode = \
    '''
    Sub Macro2()
        Range("B2:I2").Select
        Do While ActiveCell <> ""
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.Offset(1, 0).Range("A1:H1").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.Offset(1, 0).Range("A1:H1").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.Offset(1, 0).Range("A1:H1").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.Offset(1, 0).Range("A1:H1").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.Offset(1, 0).Range("A1:H1").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.Offset(1, 0).Range("A1:H1").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.Offset(1, 0).Range("A1:H1").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.Offset(-7, 1).Range("A1:H1").Select
        Loop

        Range("A2:A9").Select
        Selection.Font.Bold = True
        Do While ActiveCell <> ""
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
            ActiveCell.Offset(0, 1).Range("A1:H8").Select
        Loop
        
        Range("A10").Select
        Do While ActiveCell <> ""
            With Selection.Interior
                Selection.Font.Bold = True
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With
            ActiveCell.Offset(0, 1).Range("A1").Select
        Loop

    End Sub
    '''

    #com_instance = CreateObject("Excel.Application", dynamic = True) # USING COMTYPES
    com_instance = Dispatch("Excel.Application") # USING WIN32COM
    com_instance.Visible = True 
    com_instance.DisplayAlerts = False
    print ("Processing: %s" %scripts_dir)


    # for script_file in glob.glob(os.path.join(scripts_dir, "*.xlsx")):
    #     print ("Processing: %s" % script_file)
    #     (file_path, file_name) = os.path.split(script_file)
    #     print ("Filepath, Filename: %s %s" % (file_path,file_name))
    objworkbook = com_instance.Workbooks.Open(scripts_dir)
    #input("Enter key")
    xlmodule = objworkbook.VBProject.VBComponents.Add(1)
    xlmodule.CodeModule.AddFromString(strcode.strip())

    # run the macro
    com_instance.Application.Run('Macro2')

    # save the workbook and close
    com_instance.Workbooks(1).Close(SaveChanges=1)

    com_instance.Quit()
    return "done"

def download():
    print(fname)
    f = open("{}".format(fname),'rb')
    data2 = f.read()
    f.close()
    return fname,data2

def sendMail(to,fromEmail,pwd,sub,message):
    subject = sub
    body = message
    sender_email = fromEmail
    receiver_email = to
    password = pwd
    # Create a multipart message and set headers
    try:
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message["Bcc"] = receiver_email  # Recommended for mass emails

        # Add body to email
        message.attach(MIMEText(body, "plain"))

        filename = fname  # In same directory as script

        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email    
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        # Log in to server using secure context and send email
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, text)
        
        return "sent",filename
    
    except smtplib.SMTPAuthenticationError as e:
        return "Error",e
    
    except TypeError as er:
        return "Error",er
    
    except KeyError as err:
        return "Error",err
