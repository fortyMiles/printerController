# -*- coding: utf-8 -*-

# The file that to be an exe file
# @ Build time: 2014.11.25 17:44
# @ Author : Minchiuan Kao
# @ Last modified Time : 2014.11.26 8:10

import tornado.httpclient as httpclient
import urllib
import win32gui
import win32con
import win32api
import wmi
import time
import pickle

FILE = 'testNew.pdf'


def giveBinaryFile(fileName):
    return open(fileName, 'rb').read()


def testWriteFromBinary():
    newFile = 'new.docx'
    open(newFile, 'wb').write(giveBinaryFile(FILE))


class printSchedule:
    def __init__(self):
        self.Caption = None
        pass


class FilePrinter:
    '''
    recevie a file binary data, and build a file. Then print the file. When file is printing,
    @ Author: Minchiuan Kao
    @ Last Modified Data :2014 - 11- 25  22:03
    '''
    def __init__(self):
        self.fileBase = "c:/fileBase/"
        self.fileName = None
        self.fileID = None
        self.completedFileName = None
        self.freeJob = True  # if the free job is true, means it not have task, just free now.

    def generateAFile(self, binaryData, fileName):
        if not self.freeJob:
            print 'Error, the file printer still have remain working!'
        else:
            self.fileName = fileName
            self.completedFileName = self.fileBase + self.fileName
            try:
                open(self.fileBase + self.fileName, 'wb').write(binaryData)
                self.freeJob = False
            except:
                print 'error when generate file', self.fileName
                self.freeJob = True

    def printFile(self):
        # print file, and just print one file
        try:
            win32api.ShellExecute(None, 'print',
                                  self.completedFileName, None,
                                  '.', 0
                                  )
            self.freeJob = True
        except:
            print 'something error when print ', self.completedFileName
            self.freeJob = False


class PrinterLooker:
    '''
    Look tRhe printer's situation, and give the server information
    Current Now. Just give 3 situation : 1. Working  2. Idle 3. Error
    @ Author:  Minchiuan Kao
    @ Last Modified Date : 2014-11-25 22:03
    '''

    def __init__(self):
        self.sleepTime = 1
        self.monitor = wmi.WMI()
        self.printers = self.monitor.Win32_Printer()
        self.currentPrintedFileName = "None File"
        self.workingPrinter = None
        self.paperOut = False
        self.situation = {}
        self.printerOff = True

    def printerIsWorking(self):
        # based on the WMI package, to 'think about ' whether the printer is working now
        workingNow = False
        for printer in self.printers:
            if len(self.monitor.Win32_PrintJob(DriverName=printer.DriverName)) != 0:
                self.workingPrinter = printer
                workingNow = True
                break
        return workingNow

    def testPrinterPowerOn(self):
        # printerOff = True
        for printer in self.printers:
            if printer.EnableBIDI:
                self.printerOff = printer.WorkOffline  # working off line means if power on
                break
        print 'power off == ', self.printerOff

    def getPrintSituation(self):
        # if the printer is working, give the file name about the working now
        self.situation.clear()
        if self.printerIsWorking():
            for job in self.monitor.Win32_PrintJob(DriverName=self.workingPrinter.DriverName):
                self.currentPrintedFileName = job.Caption
                print job.Document
                self.situation['Document'] = job.Document
                print job.JobStatus
                self.situation['JobStatus'] = job.JobStatus
                if job.JobStatus and job.JobStatus.find('Paperout') >= 0:
                    self.paperOut = True
                print job.TotalPages
                self.situation['TotalPages'] = job.TotalPages
                print job.Status
                self.situation['Status'] = job.Status
        else:
            self.testPrinterPowerOn()
            if self.printerOff:
                self.situation['Status'] = 'PrinterPowerOff'
            else:
                self.situation['Status'] = 'PowerOnAndFree'

        return str(self.situation)

    def printerErrorDefineer(self):
        pass


class PrinerController:
    '''
    Do some thing that help the printer.
    Such as * close the useless windows, and analyse the windows and give the inforamtion abtou the current peoble.
    @ Author: Minchiuan Kao
    @ Last Modified Data : 2014 - 11- 25 22:49
    '''

    def __init__(self):
        self.needClosedWindows = ['HPPrinter']
        self.paperOut = False

    def enumWindowProc(self, handle, lparam):
        # give the processing action to each windows.
        title = win32gui.GetWindowText(handle)  # could get by ms-sky++ to get the fiel handle
        message  = ""
        for keyWord in self.needClosedWindows:
            if title.find(keyWord) >= 0:
                print 'find'
        #          print 'find', str(keyWord)
    #            print 'try close it'
        #           if not self.paperOut:
                print 'no paperout '*10
                win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
        #        print 'closed it'

    def closePrinterWindows(self):
        win32gui.EnumWindows(self.enumWindowProc, 0)


class PrinterClient:
    '''
    @ Descrption:To Send and Get information from Server\
                        Use Tornaod HTTP Client to completed this target
    @ Author : MinChiuan Kao
    @ Build Data: 2014- 11 - 28 09:33
    @ Last Midified Data: 2014 -11- 28 09:33
    '''
    def __init__(self, URL):
        self.http_client = httpclient.HTTPClient()
        self.URL = URL

    def getInformation(self, value):
        try:
            arg = 'command'
            value = urllib.quote(value)

            URL = '%s/?%s=%s' % (self.URL, arg, value)
            command = self.http_client.fetch(URL)
            return command.body
        except httpclient.HTTPError as e:
            # HTTPError is raised for non-200 responses; the response
            # can be found in e.response.
            print("Error: " + str(e))
        except Exception as e:
            # Other errors are possible, such as IOError.
            print("Error: " + str(e))

    def close(self):
        self.http_client.close()


class Printer:
    '''
    @ Description: Cooporated with FilePrinter, PrinterLooker, PrinerController together to build a printer client.
        # receive the information form server, and do print job, and give the information to server
    @ Author : MinChiuan Kao
    @ Build Data: 2014- 11- 25 22:30
    @ Last Modified Data: 2014-11-26 22:26
    '''
    def __init__(self):
        self.printerClient = PrinterClient('http://10.82.58.178:8000')
        self.filePrinter = FilePrinter()
        self.PrinerController = PrinerController()
        self.printerLooker = PrinterLooker()
        self.fileBinaryData = None
        self.curentFileID = None
        self.executedPrintedCmd = False  # to judge whether or not have executed the printed file command

    def connectWithServer(self):
        # send a message to server and get the replay inforamtion

        currentSituation = self.printerLooker.getPrintSituation()
        command = self.printerClient.getInformation(currentSituation)
        # command from server

        if command:
            print command
            # process inforat
        else:
            print 'No connection'

        #PROCESS COMMAND()

    def tryPrintFile(self):
##        if self.printerLoker.printerIsWorking():
##            print 'printer is working...'
##        if self.filePrinter.freeJob:
##            print 'printer has remained work'

        if self.printerLooker.printerIsWorking() == False and self.filePrinter.freeJob:
            if not self.executedPrintedCmd:
                self.filePrinter.printFile()
                self.executedPrintedCmd = True # executedPrintedCmd to be True, wait untill next self.receiveInformationFromServer to became False
                #Server need to think about the file whether been printed properly.

    def auxiliary(self,paperOut):
        # to help main printer solve some things, like close some windows.
        if paperOut:
            print 'Paper  out '* 8
        else:
            self.PrinerController.closePrinterWindows()


    def getFileFromDatabase(self):
        pass

    def run(self):
        # use multiply process to work together.
        sleepTime = 0.2
        loopNum = 0
        while True:
            self.connectWithServer()
            self.auxiliary(self.printerLooker.paperOut)
            time.sleep(sleepTime)

            loopNum += 1

def test():

    printer = Printer()

    plusInformation = '1#TestPrinterFile.pdf'
    binaryData = giveBinaryFile(FILE)

    information = (binaryData,plusInformation)
    pickle.dump(information,open('temp.cp','wb'))


    information = pickle.load( open( 'temp.cp','rb'))
   # printer.localTestIntialFile(information[0],information[1])


    # transfer by file

#    tempFile = open('temp','wb')
#    tempFile.write(information)


   # printer.localTestIntialFile(binaryData,plusInformation)
    printer.run()

if __name__ == '__main__':
    test()


