# -*- encoding: utf-8 -*-
# Version 1.0
# import utilities as excel
from os import getcwd, makedirs
from os.path import basename, exists
# from openpyxl import load_workbook
from xml.dom import minidom
from shutil import copy2
import datetime
from win32com.client import Dispatch
import urllib2
from HTMLParser import HTMLParser

class ScriptVersionHTMLParser(HTMLParser):
    def handle_data(self, data):
        if "ver" in data:
            if data[3:] != currentVersion:
                print "New updates found."
                print "Downloading new updates..."
            else:
                print "No updates found."
        return

class UpdateHTMLParser(HTMLParser):
    def handle_data(self, data):
        file = open("UPDATED.py", "w")
        content = input(data)
        file.write(content)
        file.close()
        print "Script updated from version {0} to version {1}.".format(currentVersion, data[3:])
        print "New script is called 'UPDATED.py'."
        return

def checkForUpdates():
    proxy_support = urllib2.ProxyHandler({"https": "http://iwebrolpo-1.tgn.trw.com:80"})
    opener = urllib2.build_opener(proxy_support)
    urllib2.install_opener(opener)
    html1 = urllib2.urlopen("https://github.com/EshkemPaul/ExcelScript/blob/master/README.md").read()
    html2 = urllib2.urlopen("https://raw.githubusercontent.com/EshkemPaul/ExcelScript/master/ExcelVersion1.0.py").read()
    parser1 = ScriptVersionHTMLParser()
    parser1.feed(html1)
    parser2 = UpdateHTMLParser()
    parser2.feed(html2)
    return

def createRevisionFolder():
    """
        Function creates "Revision" folder where final Excel file will be stored.
    """
    if exists(r"{0}\Revision".format(getcwd())) == False:
        makedirs(r"{0}\Revision".format(getcwd()))
    else:
        pass
    return


def createBackupFolder():
    """
        Function creates "BACKUP" folder where backup files will be stored.
    """
    if exists(r"{0}\BACKUP".format(getcwd())) == False:
        makedirs(r"{0}\BACKUP".format(getcwd()))
    else:
        pass
    return


def checkXMLfile(XMLfile):
    """
        Function checks for XML file.
        Parameters:
        - "XMLfile" - path to the XML file.
    """
    XMLfile = raw_input()
    if XMLfile.endswith(".xml"):
        if exists(XMLfile) == False:
            print "Error! File '{0}' not found! Make sure that file exists in this directory. Please type path to XML file again:".format(
                XMLfile)
            checkXMLfile(XMLfile)
        else:
            return XMLfile
    else:
        print "Error! Make sure the file is XML file. Please type path to XML file again:"
        checkXMLfile(XMLfile)
    return


def checkExcelFile(EXCELfile):
    """
        Function checks for Excel file.
        Parameters:
        - "EXCELfile" - path to the Excel file.
    """
    EXCELfile = raw_input()
    if EXCELfile.endswith(".xlsm"):
        if exists(EXCELfile) == False:
            print "Error! File '{0}' not found! Make sure that file exists in this directory. Please type path to Excel file again:".format(
                EXCELfile)
            checkExcelFile(EXCELfile)
        else:
            return EXCELfile
    else:
        print "Error! Make sure the file is Excel file with extension '*.xlsm'. Please type path to Excel file again:"
        checkExcelFile(EXCELfile)


def createBackupCopy(EXCELfile):
    """
        Create a backup copy of the original Excel file (in case something goes wrong).
        Parameters:
        - "EXCELfile" - path to the Excel file.
    """
    createBackupFolder()
    copy2(EXCELfile, r"{0}\BACKUP\COPY_{1}_{2}".format(getcwd(), timeNow, basename(EXCELfile)))
    return


# Read XML file and find tags with Test ID and Revision info and store them in lists: listID and listRev
def readXML(XMLfile):
    """
        Read XML file and find tags with Test ID and Revision info and store them in lists: listID and listRev.
        Parameters:
        - "XMLfile" - path to the XML file.
    """
    xmldom = minidom.parse(XMLfile)
    result = xmldom.getElementsByTagName("TestCase")

    for i in result:
        for j in i.getElementsByTagName("Header"):
            for k in j.getElementsByTagName("BR"):
                for l in k.getElementsByTagName("TCname"):
                    listID.append(l.firstChild.data)
                for m in k.getElementsByTagName("TCrevision"):
                    if hasattr(m.firstChild, "data"):
                        listRev.append(m.firstChild.data[9:])
                    else:
                        listRev.append("")
    return


def createExcelWithRevision(excelFile, xmlFile):
    """
        Create Excel file and add revision version info.
        Parameters:
        - "excelFile" - path to the Excel file.
        - "XMLfile" - path to the XML file.
    """

    # wb = load_workbook(filename=excelFile, use_iterators=True) # <-- if "use_iterators=True/False" not working, try using "read_only=True/False" instead
    # sheet = wb.active
    # rows = sheet.max_row
    # wb._archive.close() # if "wb._archive.close()" not working, try using "wb.close()" instead

    # doc = excel.openExcel(excelFile)
    # sheet = doc.OpenSheet(doc.GetSheetNames()[0])

    doc = Dispatch("Excel.Application")
    workbook = doc.Workbooks.Open(excelFile)
    worksheet = workbook.ActiveSheet
    rows = worksheet.UsedRange.Rows.Count

    print "Reading info from XML file..."
    readXML(xmlFile)

    print "Adding revision info to Excel file..."
    counterList = 0
    for row in range(1, rows):
        if worksheet.Cells(row, 3).Value == listID[counterList]:
            worksheet.Cells(row, 1).Value = listRev[counterList]
            counterList += 1

    try:
        createRevisionFolder()
        workbook.SaveAs(r"{0}\Revision\{1}".format(getcwd(), basename(excelFile)))
        workbook.Close(SaveChanges=1)
        doc.Quit()
    except:
        workbook.Close(SaveChanges=1)
        doc.Quit()
    del workbook
    del doc
    return

# ----------------------------------------------------------------------------#

currentVersion = "1.0"
newVersion = ""
listID = []
listRev = []

print """+-------------------+
| Add revision info |
+-------------------+\n"""

print "Checking for updates..."
checkForUpdates()

# print "Path to XML file:"
# tempXML = ""
# fileXML = checkXMLfile(tempXML)
#
# print "\nPath to Excel file with extension '*.xlsm' that you want to add revision info to:"
# tempExcel = ""
# fileExcel = checkExcelFile(tempExcel)
#
# print "\nCreating copy of Excel file in case something goes wrong..."
# timeNow = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
# createBackupCopy(fileExcel)
#
# createExcelWithRevision(fileExcel, fileXML)
#
# print "\nDone."
# print "Copy of the original Excel file is located in the same directory as script."
# print "Excel file with revisions are located in the same directory as script."
