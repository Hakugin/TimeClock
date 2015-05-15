
# _outlook.py

from wx.lib.pubsub import pub
from datetime import datetime as DT
from time import sleep
import sys
import os

import logging
log = logging.getLogger('root')

import wmi
import win32com.client as _W32C
if _W32C.gencache.is_readonly == True:
    _W32C.gencache.is_readonly = False
    _W32C.gencache.Rebuild()
    _W32C.gencache.GetGeneratePath()

__version__ = '1.0.1'

class OutlookEventHandler:
    def __init__(self):
        '''__init__ is used to append the handlers initialization to log.'''
        log.debug('Outlook Event Handler has initialized')
    def OnItemSend(self, Item=None, Cancel=None): # Email Sent
        '''Tells main module that an email was sent, using pubsub.'''
        pub.sendMessage('email.sent')
    def OnQuit(self): # Outlook closed while app is running.
        '''Tells main module that Outlook is shutting down, using pubsub.'''
        pub.sendMessage('outlook.close')

class CustomOutlookModule:
    def __init__(self):
        log.debug('Outlook module Version %s loaded.' % __version__)

    def OutlookOpenTest(self):
        '''Verify whether Outlook is already running or needs
        started.'''
        c = wmi.WMI()
        log.debug('Checking running processes for Outlook.')
        if c.Win32_Process(Name='outlook.exe'):
            log.debug('Outlook is already running.')
            obj = False
        else:
            log.debug('Outlook was not found in running processes.')
            log.debug('Starting Outlook.''')
            os.startfile('outlook')
            log.debug('Verifying that Outlook has started.')
            sleep(5)
            if c.Win32_Process(Name='outlook.exe'):
                log.debug('Outlook was found in running processes.')
                obj = True
            else:
                log.error('Outlook did not start! Exiting...')
                sys.exit(1)
        return obj

    def OnSendEmail(self, emailInfo):
        isOLOpened = self.OutlookOpenTest()
        olApp = _W32C.DispatchWithEvents(
            'outlook.application',
            OutlookEventHandler
        )

        log.debug('Creating new email for: %s' % emailInfo['Subject'])
        olMail = olApp.CreateItem(0)
        log.debug('Setting \'TO\' field: %s' % emailInfo['TO'])
        olMail.To = emailInfo['TO']
        try:
            log.debug('Setting \'CC\' field: %s' % emailInfo['CC'])
            olMail.CC = emailInfo['CC']
        except KeyError:
            log.debug('OnSendEmail did not recieve a value for \'CC\'.')
        log.debug('Setting \'Subject\' to: %s' % emailInfo['Subject'])
        olMail.Subject = emailInfo['Subject']
        log.debug('Populating email body with: %s' % emailInfo['Body'])
        olMail.Body = emailInfo['Body']
        log.debug('Sending email.')
        olMail.Send()
      # Clear up some system memory
        log.debug('Clearing system resources.')
        log.debug('Removed from Memory: %s' % olMail)
        log.debug('Removed from Memory: %s' % olApp)

    def OnSendBugReport(self, bugReport):
        log.debug('Creating bug report')
        isOLOpened = self.OutlookOpenTest()
        olApp = _W32C.DispatchWithEvents(
            'outlook.application',
            OutlookEventHandler
        )
        olMail = olApp.CreateItem(0)
        olMail.To = bugReport['TO']
        olMail.Subject = bugReport['Subject']
        olMail.Body = bugReport['Body']
        for attachment in bugReport['Attachments']:
            olMail.Attachments.Add(
                Source = attachment,
                Type = 1,
                Position = 1,
                DisplayName = os.path.basename(attachment)
            )
        log.debug('Displaying bug report to user for confirmation.')
        olMail.Display()
