
import sys
import os
import errno
import wx
from datetime import datetime as DT
from wx.lib.pubsub import pub
from glob import glob

frozen = getattr(sys, 'frozen', '')
if not frozen:
    try:
        _approot = os.path.abspath(__file__)
    except NameError:
        _approot = os.getcwd()
elif frozen in ('dll', 'console_exe', 'windows_exe'):
    _approot = os.path.abspath(sys.executable)

_approot = os.path.dirname(_approot)

sys.path.append(_approot)

# Custom Modules / Plugins
import _widgets  as _wgt
import _dialogs  as _dlg
import _logger   as _log
import _outlook  as _ol
import _database as _db

__version__ = '3.0.2'

class MainFrame(wx.Frame):
    def __init__(self, myLog, parent=None, *arg, **kw):
        self._log = myLog

        super(MainFrame, self).__init__(parent, *arg, **kw)

        self._destEmail      = ''
        self._ccEmail        = ''
        self._bugReportEmail = ''

        self.GenerateIDs()
        self.OnInitWidgets()
        self.OnInitMenu()
        self.OnInitLayout()
        self.OnBindEvents()
        self.OnLoadSettings()
        self.OnCheckDestEmail()

        self._log.debug('Creating PyTimer instance for clock.')
        self.clockTimer = wx.PyTimer(self.OnClockTimer)
        self._log.debug('Starting clock timer.')
        self.clockTimer.Start(1000)

        self._log.debug('Creating PyTimer instance for OnBtnTimer.')
        self.buttonTimer = wx.PyTimer(self.OnBtnTimer)
        self._log.debug('Starting the timer for OnBtnTimer.')
        self.buttonTimer.Start(3600000) # 1 Hour intervals

      # Outlook Module
        self._log.debug('Creating instance of the Outlook module.')
        self.olModule = _ol.CustomOutlookModule()

        pub.subscribe(self.SentEmailNotify, 'email.sent')

    def GenerateIDs(self):
        '''Generate the IDs used by user input widgets.'''
        self._idEditDstEmail = wx.NewId()
        self._log.debug('Edit Dest Email ID: %s' % self._idEditDstEmail)

        self._idEditCcEmail  = wx.NewId()
        self._log.debug('Edit CC Email ID: %s' % self._idEditCcEmail)

        self._idCcOption     = wx.NewId()
        self._log.debug('CC Option ID: %s' % self._idCcOption)

        self._idViewDstEmail = wx.NewId()
        self._log.debug('View Dest Email ID: %s' % self._idViewDstEmail)

        self._idViewCcEmail  = wx.NewId()
        self._log.debug('View CC Email ID: %s' % self._idViewCcEmail)

        self._idEditBugEmail = wx.NewId()
        self._log.debug('Edit Bug Report Email ID: %s' % self._idEditBugEmail)

        self._idViewBugEmail = wx.NewId()
        self._log.debug('View Bug Report Email ID: %s' % self._idViewBugEmail)

        self._idBugReport    = wx.NewId()
        self._log.debug('Bug Report option ID: %s' % self._idBugReport)

        self._idAbout        = wx.NewId()
        self._log.debug('About ID: %s' % self._idAbout)

        self._idDayStart     = wx.NewId()
        self._log.debug('Start of Day button ID: %s' % self._idDayStart)

        self._idLunchStart   = wx.NewId()
        self._log.debug('Start of Lunch button ID: %s' % self._idLunchStart)

        self._idLunchEnd     = wx.NewId()
        self._log.debug('End of Lunch button ID: %s' % self._idLunchEnd)

        self._idDayEnd       = wx.NewId()
        self._log.debug('End of Day button ID: %s' % self._idDayEnd)

        self._idBtnTimer     = wx.NewId()
        self._idClearTS      = wx.NewId()
        self._idSentNotify   = wx.NewId()

    def OnInitMenu(self):
      #-----------------------------------------------------------------
      # FlatMenuBar
        self._log.debug('Creating the FlatMenuBar.')
        self.menuBar = _wgt.FM.FlatMenuBar(self._panel, wx.ID_ANY, 32, 5,
            options=_wgt.FM.FM_OPT_IS_LCD)
        self.menuBar.SetBackgroundColour(wx.WHITE)
      # File Menu
        self.f_menu = _wgt.FM.FlatMenu()
        self.f_menu.Append(
            wx.ID_CLOSE,
            '&Close\tCtrl+X',
            'Exit the program',
            None
            )
        self.menuBar.Append(self.f_menu, "&File")

      # Email Settings
        self.e_menu = _wgt.FM.FlatMenu()
        self.e_menu.Append(
            self._idViewDstEmail,
            '&Destination',
            '',
        )
        self.e_menu.Append(
            self._idViewCcEmail,
            '&CC Email',
            '',
        )
        self.e_menu.Append(
            self._idViewBugEmail,
            '&Bug Report Email',
            '',
        )
        self.menuBar.Append(self.e_menu, '&View')

      # Settings Menu
        self.s_menu = _wgt.FM.FlatMenu()
        self.s_menu.AppendCheckItem(
            self._idCcOption,
            '&Include CC',
            'Include CC email address.'
            )
        self.s_menu.AppendCheckItem(
            self._idSentNotify,
            'Email Sent &Notification',
            'Enable / Disable Notification'
            )
        self.s_menu.Append(
            self._idClearTS,
            '&Clear Locks',
            'Clear button lock-outs.',
            None
            )

        emailSubMod = _wgt.FM.FlatMenu()
        editDstEmail = _wgt.FM.FlatMenuItem(
            emailSubMod,
            self._idEditDstEmail,
            '&Destination',
            '',
            wx.ITEM_NORMAL
        )
        emailSubMod.AppendItem(editDstEmail)
        editCcEmail = _wgt.FM.FlatMenuItem(
            emailSubMod,
            self._idEditCcEmail,
            '&CC Email',
            '',
            wx.ITEM_NORMAL
        )
        emailSubMod.AppendItem(editCcEmail)
        editBugEmail = _wgt.FM.FlatMenuItem(
            emailSubMod,
            self._idEditBugEmail,
            '&Bug Report Email',
            '',
            wx.ITEM_NORMAL
        )
        emailSubMod.AppendItem(editBugEmail)
        self.s_menu.AppendMenu(-1, '&Edit Addresses', emailSubMod)
        self.menuBar.Append(self.s_menu, '&Settings')

      # Help Menu
        self.h_menu = _wgt.FM.FlatMenu()
        self.h_menu.Append(
            self._idBugReport,
            '&Report Bug',
            '',
        )
        self.h_menu.Append(
            self._idAbout,
            '&About',
            '',
        )

#        self._idOptionTest1 = wx.NewId()
#        self.h_menu.Append(self._idOptionTest1, '&Save', '', None)

        self.menuBar.Append(self.h_menu, '&Help')

    def OnInitWidgets(self):
        self._log.debug('Initializing GUI Widgets.')
        self._log.debug('Creating the panel.')
        self._panel = wx.Panel(self)
        self._panel.SetBackgroundColour(wx.Colour(220,220,220,255))

        self._log.debug('Creating the Title static text.')
        self._titleBox = wx.StaticText(
            self._panel, -1, 'Clock Puncher %s' % __version__,
            style=wx.NO_BORDER|wx.TE_CENTER
        )
        self._titleBox.SetBackgroundColour(wx.WHITE)

      #-----------------------------------------------------------------
      # Buttons and TextCtrls
        self._log.debug('Creating the Buttons and TextCtrls.')
        self.btnStartDay = _wgt.CustomPB(
                                self._panel,
                                id=self._idDayStart,
                                label='Start of Day',
                                style=wx.SIMPLE_BORDER,
                                size=(93,-1)
                                )
        self.txtStartDay = _wgt.CustomTextCtrl(
                                self._panel,
                                wx.ID_ANY,
                                value='',
                                style=wx.NO_BORDER,
                                size=(125,-1)
                            )

        self.btnStartLunch = _wgt.CustomPB(
                                self._panel,
                                id=self._idLunchStart,
                                label='Start of Lunch',
                                style=wx.SIMPLE_BORDER,
                                size=(93,-1)
                                )
        self.txtStartLunch = _wgt.CustomTextCtrl(
                                self._panel,
                                wx.ID_ANY,
                                value='',
                                style=wx.NO_BORDER,
                                size=(125,-1)
                            )

        self.btnEndLunch = _wgt.CustomPB(
                                self._panel,
                                id=self._idLunchEnd,
                                label='End of Lunch',
                                style=wx.SIMPLE_BORDER,
                                size=(93,-1)
                                )
        self.txtEndLunch = _wgt.CustomTextCtrl(
                                self._panel,
                                wx.ID_ANY,
                                value='',
                                style=wx.NO_BORDER,
                                size=(125,-1)
                            )

        self.btnEndDay = _wgt.CustomPB(
                                self._panel,
                                id=self._idDayEnd,
                                label='End of Day',
                                style=wx.SIMPLE_BORDER,
                                size=(93,-1)
                                )
        self.txtEndDay = _wgt.CustomTextCtrl(
                                self._panel,
                                wx.ID_ANY,
                                value='',
                                style=wx.NO_BORDER,
                                size=(125,-1)
                            )

        self._log.debug('Creating the Clock TextCtrl.')
        self.stTxtClock = _wgt.CustomTextCtrl(
                                self._panel,
                                -1,
                                'Starting Clock..',
                                style=wx.NO_BORDER|wx.TE_CENTER
                            )
        self.stTxtClock.SetBackgroundColour(wx.WHITE)

        self._log.debug('Creating the list for the mouse event binds.')
        self._mouseList = (self._panel, self._titleBox,
                            self.txtStartDay, self.txtStartLunch,
                            self.txtEndLunch, self.txtEndDay,
                            self.stTxtClock
                        )

    def OnInitLayout(self):
        '''Configure Layout.'''
        self._log.debug('Initializing the layout.')
        mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        panelSizer = wx.BoxSizer(wx.VERTICAL)
        titleSizer = wx.BoxSizer(wx.HORIZONTAL)
        menuSizer = wx.BoxSizer(wx.HORIZONTAL)

        titleSizer.Add(self._titleBox, 1, wx.ALL|wx.EXPAND|wx.CENTER,1)

        menuSizer.Add(self.menuBar, 1, wx.EXPAND, 1)

        widgetSizer = wx.FlexGridSizer(rows=4, cols=2, vgap=1, hgap=2)
        widgetSizer.AddMany(
            [
                (self.btnStartDay,   1, wx.ALL|wx.EXPAND, 3),
                (self.txtStartDay,   1, wx.ALL|wx.EXPAND, 3),
                (self.btnStartLunch,   1, wx.ALL|wx.EXPAND, 3),
                (self.txtStartLunch,   1, wx.ALL|wx.EXPAND, 3),
                (self.btnEndLunch, 1, wx.ALL|wx.EXPAND, 3),
                (self.txtEndLunch, 1, wx.ALL|wx.EXPAND, 3),
                (self.btnEndDay,  1, wx.ALL|wx.EXPAND, 3),
                (self.txtEndDay,  1, wx.ALL|wx.EXPAND, 3)
            ]
        )

        clockSizer = wx.BoxSizer(wx.HORIZONTAL)
        clockSizer.Add(self.stTxtClock, 1, wx.ALL|wx.EXPAND|wx.CENTER, 1)

        panelSizer.Add(titleSizer, 0, wx.ALL|wx.EXPAND, 2)
        panelSizer.Add(menuSizer, 0, wx.ALL|wx.EXPAND, 3)
        panelSizer.Add(widgetSizer, 3, wx.EXPAND,1)
        panelSizer.Add(clockSizer, 0, wx.ALL|wx.EXPAND|wx.CENTER, 2)
        self._panel.SetSizer(panelSizer)

        mainSizer.Add(self._panel, 1, wx.ALL|wx.EXPAND, 2)
        self.SetSizerAndFit(mainSizer)
        self.SetPosition( (400,300) )
        self.Layout()

    def OnBindEvents(self):
        self._log.debug('Binding events and widgets to functions.')
        self.Bind(wx.EVT_MOTION, self.OnMouseMove)
        self.Bind(wx.EVT_PAINT, self.OnPaint)

        for _widget in self._mouseList:
            _widget.Bind(wx.EVT_MOTION, self.OnMouseMove)

        self.Bind(wx.EVT_MENU, self.OnClose, id=wx.ID_CLOSE)
        self.Bind(wx.EVT_MENU, self.OnClearLocks, id=self._idClearTS)
        self.Bind(wx.EVT_MENU, self.OnEditEmail, id=self._idEditDstEmail)
        self.Bind(wx.EVT_MENU, self.OnEditEmail, id=self._idEditCcEmail)
        self.Bind(wx.EVT_MENU, self.OnEditEmail, id=self._idEditBugEmail)
        self.Bind(wx.EVT_MENU, self.OnViewEmail, id=self._idViewDstEmail)
        self.Bind(wx.EVT_MENU, self.OnViewEmail, id=self._idViewCcEmail)
        self.Bind(wx.EVT_MENU, self.OnViewEmail, id=self._idViewBugEmail)
        self.Bind(wx.EVT_MENU, self.OnAboutDlg, id=self._idAbout)

        self.Bind(wx.EVT_BUTTON, self.OnDayStart, id=self._idDayStart)
        self.Bind(wx.EVT_BUTTON, self.OnLunchStart, id=self._idLunchStart)
        self.Bind(wx.EVT_BUTTON, self.OnLunchEnd, id=self._idLunchEnd)
        self.Bind(wx.EVT_BUTTON, self.OnDayEnd, id=self._idDayEnd)

        self.Bind(wx.EVT_MENU, self.OnSaveSettings, id=self._idCcOption)
        self.Bind(wx.EVT_MENU, self.OnCheckCCEmail, id=self._idCcOption)
        self.Bind(wx.EVT_MENU, self.OnSaveSettings, id=self._idSentNotify)
        self.Bind(wx.EVT_MENU, self.OnReportBug, id=self._idBugReport)

    def OnSaveSettings(self, event=None):
        '''Needs Implemented'''
        self._log.debug('Saving Settings.')
        ccOpt = self.menuBar.FindMenuItem(self._idCcOption)
        sentNotifyOpt = self.menuBar.FindMenuItem(self._idSentNotify)

        cfgSettings = {}
        cfgSettings['dest_email']        = self._destEmail
        cfgSettings['cc_email']          = self._ccEmail
        cfgSettings['cc_option']         = ccOpt.IsChecked()
        cfgSettings['email_notify']      = sentNotifyOpt.IsChecked()
        cfgSettings['bug_report_dest']   = self._bugReportEmail

        _db.SaveSettings(cfgSettings)

    def OnLoadSettings(self, event=None):
        '''Needs Implemented'''
        ccOpt = self.menuBar.FindMenuItem(self._idCcOption)
        notifyOpt = self.menuBar.FindMenuItem(self._idSentNotify)
        prgSettings = _db.LoadSettings()
        for key, val in prgSettings.items():
            if key == 'dest_email':
                self._destEmail = val

            if key == 'cc_email':
                self._ccEmail = val

            if key == 'bug_report_dest':
                self._bugReportEmail = val

            if key == 'cc_option':
                if int(val) == 1:
                    ccOpt.Check(True)
                else:
                    ccOpt.Check(False)

            if key == 'email_notify':
                if int(val) == 1:
                    notifyOpt.Check(True)
                else:
                    notifyOpt.Check(False)

        curDate = DT.strftime(DT.now(), '%Y-%m-%d')
        timeStamps = _db.OnLoadTimeStamps(curDate)
        for key, val in timeStamps.items():
            if key == 'dayS':
                self.txtStartDay.SetValue(val)
                if val != '':
                    self.btnStartDay.Disable()

            if key == 'lunchS':
                self.txtStartLunch.SetValue(val)
                if val != '':
                    self.btnStartLunch.Disable()

            if key == 'lunchE':
                self.txtEndLunch.SetValue(val)
                if val != '':
                    self.btnEndLunch.Disable()

            if key == 'dayE':
                self.txtEndDay.SetValue(val)
                if val != '':
                    self.btnEndDay.Disable()

    def OnCheckDestEmail(self):
        if self._destEmail == '':
            self._log.debug('Destination email address not found in settings!')
            self._log.debug('Prompting user for destination email address.')
            dlg = _dlg.EnterEmail(self, eType='DEST',pos=(400,300))
            prompt = _dlg.MyMessageDialog(self,
                        msg = 'You must enter a\'Destination\' email address.',
                        myStyle = None
                        )
            prompt.OnSetBorderColour(wx.RED)
            while True:
                if dlg.ShowModal() == wx.ID_OK:
                    self._destEmail = dlg.GetAddress()
                    if self._destEmail != '':
                        prompt.Destroy()
                        dlg.Destroy()
                        break
                    else:
                        prompt.ShowModal()
                else: 
                    prompt.ShowModal()

    def OnCheckCCEmail(self, event=None):
        ccOpt = self.menuBar.FindMenuItem(self._idCcOption)
        if self._ccEmail == '':
            self._log.debug('CC Email address not found.')
            self._log.debug('Prompting user.')
            dlg = _dlg.EnterEmail(self, eType='CC')
            prompt = _dlg.MyMessageDialog(self,
                        msg = 'You must enter an email address to \'Carbon Copy\'.',
                        myStyle = None
                        )
            prompt.OnSetBorderColour(wx.RED)
            while True:
                if dlg.ShowModal() == wx.ID_OK:
                    self._ccEmail = dlg.GetAddress()
                    if self._ccEmail != '':
                        prompt.Destroy()
                        dlg.Destroy()
                        break
                    else:
                        prompt.ShowModal()
                else:
                    prompt.ShowModal()

    def OnAboutDlg(self, event):
        '''Needs Implemented'''
        self._log.debug('OnAboutDlg triggered.')
        dlg = _dlg.AboutWindow(self, ver = __version__)
        dlg.Show()

    def OnViewEmail(self, event):
        '''View saved Dest / CC email addresses.'''
        self._log.debug('OnViewEmail triggered.')
        evtID = event.GetId()
        if evtID == self._idViewDstEmail:
            dlg = _dlg.ViewEmail(self, eType='DEST', eMail=self._destEmail)
        elif evtID == self._idViewCcEmail:
            dlg = _dlg.ViewEmail(self, eType='CC', eMail=self._ccEmail)
        elif evtID == self._idViewBugEmail:
            dlg = _dlg.ViewEmail(self, eType='BUG', eMail=self._bugReportEmail)
        dlg.ShowModal()

    def OnEditEmail(self, event):
        '''Edit saved Dest / CC email addresses.'''
        self._log.debug('OnEditEmail triggered.')
        evtID = event.GetId()
        if evtID == self._idEditDstEmail:
            self._log.debug('Editing Destination Email.')
            dlg = _dlg.EnterEmail(self, eType='DEST')
        elif evtID == self._idEditCcEmail:
            dlg = _dlg.EnterEmail(self, eType='CC')
            self._log.debug('Editing CC Email.')
        elif evtID == self._idEditBugEmail:
            dlg = _dlg.EnterEmail(self, eType='BUG')
            self._log.debug('Editing Bug Report Email.')

        if dlg.ShowModal() == wx.ID_OK:
            emailAddress = dlg.GetAddress()
            if evtID == self._idEditDstEmail:
                self._log.debug(
                    'Destination email changed to: %s' % emailAddress
                )
                self._destEmail = emailAddress
            elif evtID == self._idEditCcEmail:
                self._log.debug(
                    'CC email changed to: %s' % emailAddress
                )
                self._ccEmail = emailAddress
            elif evtID == self._idEditBugEmail:
                self._log.debug(
                    'Bug Report email changed to: %s' % emailAddress
                )
                self._bugReportEmail = emailAddress
        self.OnSaveSettings()
        dlg.Destroy()

    def OnClearLocks(self, event):
        '''Re-Enable buttons.'''
        self._log.debug('OnClearLocks triggered.')
        dlg = _dlg.ClearBtnLocks(self)
        dlg.OnSetBorderColour(wx.Colour(46,139,87,255))
        if dlg.ShowModal() == (wx.ID_OK):
            self.OnSaveSettings()
        dlg.Destroy()

    def SentEmailNotify(self, event=None):
        '''Create pop-up when Outlook successfully sends mail.'''
        self._log.debug('SentEmailNotify triggered.')
        sentNotifyOpt = self.menuBar.FindMenuItem(self._idSentNotify)
        if sentNotifyOpt.IsChecked():
            msg = ("Your email has been sent by Outlook.\
            \nThis can be verified by checking your sent folder.")
            dlg = _dlg.MyMessageDialog(self, msg, myStyle=None)
            dlg.OnSetBorderColour(wx.Colour(195,82,82,255))
            dlg.Show()

    def OnDayStart(self, event):
        '''Send "Start of Day" email.'''
        self._log.debug('OnDayStart triggered.')
        timestamp = DT.now()
        tFormat = "%m/%d/%y %I:%M:%S %p"
        evtObj = event.GetEventObject()
        if event.GetId() != self._idDayStart:
            self._log.critical(
                'OnDayStart function received event from \'%s\''%(
                    evtObj.GetLabel()
                    )
                )
            return
        self._log.debug('OnDayStart received event ID: %s' % event.GetId())
        ccOpt = self.menuBar.FindMenuItem(self._idCcOption)
        curDate = DT.strftime(timestamp, '%Y-%m-%d')
        curTime = DT.strftime(timestamp, tFormat)
        emailInfo = {}

        if ccOpt.IsChecked() and self._ccEmail != (
                '' and 'Invalid' and []):
            emailInfo['CC'] = self._ccEmail

        emailInfo['TO'] = self._destEmail
        emailInfo['Subject'] = 'Start of Day'
        emailInfo['Body'] = 'Email Creation: %s' % timestamp
        self.txtStartDay.SetValue(DT.strftime(timestamp, tFormat))
        self.olModule.OnSendEmail(emailInfo)
        _db.OnUpdateTimeStamp(theDate=curDate, dayS=curTime)
        evtObj.Disable()
        event.Skip()

    def OnLunchStart(self, event):
        '''Send "Start of Lunch" email.'''
        self._log.debug('OnLunchStart triggered.')
        timestamp = DT.now()
        tFormat = "%m/%d/%y %I:%M:%S %p"
        evtObj = event.GetEventObject()
        if event.GetId() != self._idLunchStart:
            self._log.critical(
                'OnLunchStart function received event from \'%s\''%(
                    evtObj.GetLabel()
                    )
                )
            return
        self._log.debug('OnLunchStart received event ID: %s' % event.GetId())
        ccOpt = self.menuBar.FindMenuItem(self._idCcOption)
        curDate = DT.strftime(timestamp, '%Y-%m-%d')
        curTime = DT.strftime(timestamp, tFormat)
        emailInfo = {}

        if ccOpt.IsChecked() and self._ccEmail != (
                '' and 'Invalid' and []):
            emailInfo['CC'] = self._ccEmail

        emailInfo['TO'] = self._destEmail
        emailInfo['Subject'] = 'Start of Lunch'
        emailInfo['Body'] = 'Email Creation: %s' % timestamp
        self.txtStartLunch.SetValue(DT.strftime(timestamp, tFormat))
        self.olModule.OnSendEmail(emailInfo)
        _db.OnUpdateTimeStamp(theDate=curDate, lunchS=curTime)
        evtObj.Disable()
        event.Skip()

    def OnLunchEnd(self, event):
        '''Send "End of Lunch" email.'''
        self._log.debug('OnLunchEnd triggered.')
        timestamp = DT.now()
        tFormat = "%m/%d/%y %I:%M:%S %p"
        evtObj = event.GetEventObject()
        if event.GetId() != self._idLunchEnd:
            self._log.critical(
                'OnLunchEnd function received event from \'%s\''%(
                    evtObj.GetLabel()
                    )
                )
            return
        self._log.debug('OnLunchEnd received event ID: %s' % event.GetId())
        ccOpt = self.menuBar.FindMenuItem(self._idCcOption)
        curDate = DT.strftime(timestamp, '%Y-%m-%d')
        curTime = DT.strftime(timestamp, tFormat)
        emailInfo = {}

        if ccOpt.IsChecked() and self._ccEmail != (
                '' and 'Invalid' and []):
            emailInfo['CC'] = self._ccEmail

        emailInfo['TO'] = self._destEmail
        emailInfo['Subject'] = 'End of Lunch'
        emailInfo['Body'] = 'Email Creation: %s' % timestamp
        self.txtEndLunch.SetValue(DT.strftime(timestamp, tFormat))
        self.olModule.OnSendEmail(emailInfo)
        _db.OnUpdateTimeStamp(theDate=curDate, lunchE=curTime)
        evtObj.Disable()
        event.Skip()

    def OnDayEnd(self, event):
        '''Send "End of Day" email.'''
        self._log.debug('OnDayEnd triggered.')
        timestamp = DT.now()
        tFormat = "%m/%d/%y %I:%M:%S %p"
        evtObj = event.GetEventObject()
        if event.GetId() != self._idDayEnd:
            self._log.critical(
                'OnDayEnd function received event from \'%s\''%(
                    evtObj.GetLabel()
                    )
                )
            return
        self._log.debug('OnDayEnd received event ID: %s' % event.GetId())
        ccOpt = self.menuBar.FindMenuItem(self._idCcOption)
        curDate = DT.strftime(timestamp, '%Y-%m-%d')
        curTime = DT.strftime(timestamp, tFormat)
        emailInfo = {}

        if ccOpt.IsChecked() and self._ccEmail != (
                '' and 'Invalid' and []):
            emailInfo['CC'] = self._ccEmail

        emailInfo['TO'] = self._destEmail
        emailInfo['Subject'] = 'End of Day'
        emailInfo['Body'] = 'Email Creation: %s' % timestamp
        self.txtEndDay.SetValue(DT.strftime(timestamp, tFormat))
        self.olModule.OnSendEmail(emailInfo)
        _db.OnUpdateTimeStamp(theDate=curDate, dayE=curTime)
        evtObj.Disable()
        event.Skip()

    def OnReportBug(self, event):
        '''Generate and send bug report.'''
        dlg = _dlg.BugReport(self)
        if dlg.ShowModal() == wx.ID_OK:
            tFormat = "%Y/%m/%d %H:%M:%S"
            timestamp = DT.strftime(DT.now(), tFormat)
            emailInfo = {}
            emailInfo['TO'] = self._bugReportEmail
            emailInfo['Subject'] = 'BUG REPORT: %s' % timestamp
            emailInfo['Body'] = dlg.OnGetSummary()+'\n\nCreated: %s' % timestamp
            #print emailInfo
            with open('debug_report.log','w') as logOutput:
                for entry in _db.ReadLogData():
                    logOutput.write(str(entry)+'\n')
            emailInfo['Attachments'] = glob(
                os.path.join(_approot,'*.log')
            )
            self.olModule.OnSendBugReport(emailInfo)
        dlg.Destroy()
        for attached in emailInfo['Attachments']:
            self.OnSilentRemoveFile(attached)

    def OnSilentRemoveFile(self, filename):
        try:
            os.remove(filename)
        except OSError as e:
            if e.errno != errno.ENOENT:
                self._log.error('%s' % e)

    def OnPaint(self, event):
        evtObj = event.GetEventObject()
        evtObjBG = evtObj.GetBackgroundColour()

        dc = wx.PaintDC(evtObj)
        dc = wx.GCDC(dc)
        w, h = self.GetSizeTuple()
        r = 10
        dc.SetPen( wx.Pen("#000000",3) )
        dc.SetBrush( wx.Brush(evtObjBG))
        dc.DrawRectangle( 0,0,w,h )

    def OnMouseMove(self, event):
        """implement dragging"""
        if not event.Dragging():
            self._dragPos = None
            return
        #self.CaptureMouse()
        if not self._dragPos:
            self._dragPos = event.GetPosition()
        else:
            pos = event.GetPosition()
            displacement = self._dragPos - pos
            self.SetPosition( self.GetPosition() - displacement )
        #self.ReleaseMouse()
        event.Skip()

    def OnBtnTimer(self):
        '''Re-Enable buttons if date has changed and program is still running.'''
        date_change = False
        if DT.now().strftime('%m/%d/%y') != (
                self.txtStartDay.GetValue()[:-12]):
            self._log.debug(
                'Timestamp date does not match current date, enabling \'Start of Day\' button.')
            self.btnDayStart.Enable()
            if not date_change:
                date_change = True
                self.SaveTimestamps()

        if DT.now().strftime('%m/%d/%y') != (
                self.txtStartLunch.GetValue()[:-12]):
            self._log.debug(
                'Timestamp date does not match current date, enabling \'Start of Lunch\' button.')
            self.btnLunchStart.Enable()
            if not date_change:
                date_change = True
                self.SaveTimestamps()

        if DT.now().strftime('%m/%d/%y') != (
                self.txtEndLunch.GetValue()[:-12]):
            self._log.debug(
                'Timestamp date does not match current date, enabling \'End of Lunch\' button.')
            self.btnLunchEnd.Enable()
            if not date_change:
                date_change = True
                self.SaveTimestamps()

        if DT.now().strftime('%m/%d/%y') != (
                self.txtEndDay.GetValue()[:-12]):
            self._log.debug(
                'Timestamp date does not match current date, enabling \'End of Day\' button.')
            self.btnDayEnd.Enable()
            if not date_change:
                date_change = True
                self.SaveTimestamps()

    def OnClockTimer(self):
        '''Update Clock at the bottom of the program.'''
        ts = DT.now().strftime('%Y/%m/%d - %I:%M:%S %p')
        self.stTxtClock.SetLabel(ts)

    def SaveTimestamps(self):
        '''Save timestamps when program closes.'''
        self._log.debug('Saving Timestamps.')

        dayStart   = self.txtStartDay.GetValue()
        lunchStart = self.txtStartLunch.GetValue()
        lunchEnd   = self.txtEndLunch.GetValue()
        dayEnd     = self.txtEndDay.GetValue()

        curDate = DT.strftime(DT.now(), '%m/%d/%y')
        if curDate != dayStart[:-12]:
            dayStart = ''

        if curDate != lunchStart[:-12]:
            lunchStart = ''

        if curDate != lunchEnd[:-12]:
            lunchEnd = ''

        if curDate != dayEnd[:-12]:
            dayEnd = ''

        curDate = DT.strftime(DT.now(), '%Y-%m-%d')
        _db.OnUpdateTimeStamp(
            theDate = curDate,
            dayS    = dayStart,
            lunchS  = lunchStart,
            lunchE  = lunchEnd,
            dayE    = dayEnd
            )

    def OnClose(self, event):
        '''I believe this is self explanatory.'''
        self._log.debug('OnClose triggered. Closing program.')
        self.clockTimer.Stop()
        self.buttonTimer.Stop()

        self.OnSaveSettings()
        self.SaveTimestamps()
        self.Close(force=True)


def RunApp():
    app = wx.App()

    log = _log.logging.getLogger('root')
    log.addHandler(_log.SQLAlchemyHandler())

    if len(sys.argv) > 1:
        if str(sys.argv[1]).lower() == '--debug':
            log.setLevel('DEBUG')
            log.debug('Logging level set to \'DEBUG\'.')
        else:
            print('')
            print('-' * 40)
            print('The only option available is \'--debug\'.')
            print('Which will increase the log verbosity.')
            print('-' * 40)
    else:
        log.setLevel('WARNING')    
        log.debug('Logging level set to \'WARNING\'')

    frame = MainFrame(title='Clock Puncher',
        style = ( wx.CLIP_CHILDREN | wx.NO_BORDER ),
        myLog = log
    )
    frame.Show()

    app.MainLoop()

if __name__ == '__main__':
    RunApp()
