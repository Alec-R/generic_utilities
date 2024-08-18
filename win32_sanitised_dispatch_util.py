from PyPDF2 import PdfFileReader

# OS requirements
import win32com.client as win32
import glob
import pathlib
import shutil
import os

#########
# outlook COM email util
#########

def dispatch(app_name:str):
    """ dynamic dispatch for COM """
    try: # this should work but might fail because of gen_py cache
        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    except AttributeError:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))

        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    return app



def email_template(**kwargs):
    """
    takes x and creates an email
    """
    to = kwargs.get('to','')
    subject = kwargs.get('subject','')
    body = "Hi \n"

    filename = 'KID.pdf'
    folder = ' - '
    p1 = folder + "*"

    fullpath = glob.glob(p1)[0] + '/'
    properpath = pathlib.Path(fullpath)
    p2 = properpath/filename

    newname = "_1"
    p2 = copy_save(p1,filename,newname,'C:/temp')

    attach = str(p2)

    makeEmail(to, subject, body, attach)
    os.remove(p2)



def makeEmail(**kwargs):
    """ creates an email with the passed parameters - case sensitive

    subject: subject of the email
    to: recipient
    cc: copy
    bcc: blind copy
    body: content, must be a string.
    """
    # Open up an outlook email
    outlook = dispatch('Outlook.Application')
    new_mail = outlook.CreateItem(0)
    
    # email dispatch contents
    new_mail.Subject = kwargs.get('subject')
    new_mail.To = kwargs.get('to')
    new_mail.CC = kwargs.get('cc')
    new_mail.BCC = kwargs.get('bcc')
    new_mail.Body = kwargs.get('body')
    # will probably not be used, but... new_mail.SentOnBehalfOfName = behalf

    # The file needs to be a string not a p1 object
    if kwargs.get('attach'): new_mail.Attachments.Add(Source=kwargs.get('attach')) 

    # Display the email
    new_mail.Display(False)
    new_mail.Save()


#########
# file management, metadata...
#########

def copy_save(folderpath,filename,newname,outputdir) -> pathlib.Path:
    """ copy a file from a glob matched path to a new location, return the path of the copy destination"""
    p1 = folderpath + "*"
    fullpath = glob.glob(p1)[0] + '/'
    properpath = pathlib.Path(fullpath)

    if properpath.exists():
        f1 = properpath / filename
        fextension = f1.suffix
        outputpath = pathlib.Path(outputdir)
        new_name = f'{newname}{fextension}'
        p2 = outputpath / new_name
        shutil.copy(f1, p2)
        return p2
    else:
        raise ValueError('Path not found')

def pdf_metadata(fullpath):
    with open(fullpath, 'rb') as f:
        pdf_toread = PdfFileReader(f)
        pdf_info = pdf_toread.getDocumentInfo()
    return dict(pdf_info)

# END
