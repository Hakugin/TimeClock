
# _dialogs.py

import wx
import _widgets as _wgt
from wx.lib.wordwrap import wordwrap

__version__ = '1.0.0'

import logging
log = logging.getLogger('root')

class CustomDialog(wx.Dialog):
    def __init__(self, parent, *arg, **kw):
        style = (wx.NO_BORDER | wx.CLIP_CHILDREN)
        self.borderColour = wx.BLACK
        wx.Dialog.__init__(self, parent, title='Fancy', style = style)
        self.Bind(wx.EVT_MOTION, self.OnMouse)
        self.Bind(wx.EVT_PAINT, self.OnPaint)

    def OnSetBorderColour(self, color):
        self.borderColour = color

    def OnPaint(self, event):
        evtObj = event.GetEventObject()
        evtObjBG = evtObj.GetBackgroundColour()

        dc = wx.PaintDC(evtObj)
        dc = wx.GCDC(dc)
        w, h = self.GetSizeTuple()
        r = 10
        dc.SetPen( wx.Pen(self.borderColour,3) )
        dc.SetBrush( wx.Brush(evtObjBG))
        dc.DrawRectangle( 0,0,w,h )

    def OnMouse(self, event):
        """implement dragging"""
        if not event.Dragging():
            self._dragPos = None
            return
        self.CaptureMouse()
        if not self._dragPos:
            self._dragPos = event.GetPosition()
        else:
            pos = event.GetPosition()
            displacement = self._dragPos - pos
            self.SetPosition( self.GetPosition() - displacement )
        self.ReleaseMouse()
        event.Skip()


class MyMessageDialog(CustomDialog):
    def __init__(self, parent, msg, myStyle, *arg, **kw):
        self._msg = msg
        self._style = myStyle
        super(MyMessageDialog, self).__init__(parent, *arg, **kw)

        self._panel = wx.Panel(self)
        self._message = wordwrap(str(self._msg), 350, wx.ClientDC(self))
        self.stcText = wx.StaticText(self._panel, -1, self._message)
        self.stcText.SetFocus()
        self.btnOk = _wgt.CustomPB(self._panel, id=wx.ID_OK, label='OK')

        mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        panelSizer = wx.BoxSizer(wx.VERTICAL)
        panelSizer.Add(self.stcText, 1, wx.ALL|wx.EXPAND|wx.CENTRE, 2)
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        btnSizer.Add(self.btnOk, 0, wx.ALL|wx.CENTRE, 5)

        if self._style != None and 'INC_CANCEL' in self._style:
            self.btnCancel = _wgt.CustomPB(self._panel,
                                id=wx.ID_CANCEL, label='Cancel')
            btnSizer.Add(self.btnCancel, 0, wx.ALL|wx.CENTRE, 5)

        panelSizer.Add(btnSizer, 0, wx.ALL|wx.CENTRE, 5)
        self._panel.SetSizer(panelSizer)
        mainSizer.Add(self._panel, 1, wx.ALL|wx.EXPAND|wx.CENTRE, 2)
        self.SetSizerAndFit(mainSizer)

        self._panel.Bind(wx.EVT_MOTION, self.OnMouse)
        self.Bind(wx.EVT_BUTTON, self.OnOk, id=wx.ID_OK)

        self.CenterOnParent()

    def OnOk(self, event):
        self.EndModal(wx.ID_OK)


class ClearBtnLocks(CustomDialog):
    def __init__(self, parent, *arg, **kw):
        super(ClearBtnLocks, self).__init__(parent, *arg, **kw)
        self.parent = parent

        self.OnInitUI()
        self.OnInitLayout()
        self.OnBindEvents()
        self.CenterOnParent()
        self.OnSetBorderColour(wx.Colour(46,139,87,255))

    def OnInitUI(self):
        self._panel = wx.Panel(self)
        
        self.cbDayStart = wx.CheckBox(self._panel, -1, 'Clear Start of Day')
        self.cbLunchStart = wx.CheckBox(self._panel, -1, 'Clear Start of Lunch')
        self.cbLunchEnd = wx.CheckBox(self._panel, -1, 'Clear End of Lunch')
        self.cbDayEnd = wx.CheckBox(self._panel, -1, 'Clear End of Day')
        self.cbCheckAll = wx.CheckBox(self._panel, 0, 'Select All')

        self.cbList = ( self.cbDayStart, self.cbLunchStart,
            self.cbLunchEnd, self.cbDayEnd )
        
        self.btnOk = _wgt.CustomPB(self._panel, id=wx.ID_OK, label='OK')
        self.btnCancel = _wgt.CustomPB(self._panel, id=wx.ID_CANCEL, label='Cancel')

    def OnInitLayout(self):
        mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        panelSizer = wx.BoxSizer(wx.VERTICAL)

        statBox = wx.StaticBox(self._panel, label='Which Buttons to Enable?')
        sbSizer = wx.StaticBoxSizer(statBox, wx.VERTICAL)

        cbSizer = wx.FlexGridSizer(3, 2, 3, 3)
        cbSizer.AddMany(
            [
                (self.cbDayStart, 0, wx.ALL, 1),
                (self.cbLunchStart, 0, wx.ALL, 1),
                (self.cbLunchEnd, 0, wx.ALL, 1),
                (self.cbDayEnd, 0, wx.ALL, 1),
                (self.cbCheckAll, 0, wx.ALL, 1)
            ]
        )
        sbSizer.Add(cbSizer, 0, wx.ALL, 1)

        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        btnSizer.Add(self.btnOk, 0, wx.ALL, 5)
        btnSizer.Add(self.btnCancel, 0, wx.ALL, 5)
        sbSizer.Add(btnSizer, 1, wx.ALL|wx.CENTRE, 1)

        panelSizer.Add(sbSizer, 0, wx.ALL | wx.EXPAND, 5)
        self._panel.SetSizer(panelSizer)

        mainSizer.Add(self._panel, 1, wx.ALL|wx.EXPAND|wx.CENTRE, 2)
        self.SetSizerAndFit(mainSizer)

    def OnBindEvents(self):
        self._panel.Bind(wx.EVT_MOTION, self.OnMouse)
        self.Bind(wx.EVT_BUTTON, self.OnOk, id=wx.ID_OK)
        for cbOption in self.cbList:
            cbOption.Bind(wx.EVT_CHECKBOX, self.ClearSelectAll)
        self.cbCheckAll.Bind(wx.EVT_CHECKBOX, self.OnSelectAll)

    def OnSelectAll(self, event):
        if self.cbCheckAll.GetValue():
            self.cbDayStart.SetValue(True)
            self.cbLunchStart.SetValue(True)
            self.cbLunchEnd.SetValue(True)
            self.cbDayEnd.SetValue(True)
        else:
            self.cbDayStart.SetValue(False)
            self.cbLunchStart.SetValue(False)
            self.cbLunchEnd.SetValue(False)
            self.cbDayEnd.SetValue(False)

    def ClearSelectAll(self, event):
        self.cbCheckAll.SetValue(False)

    def OnOk(self, event):
        log.debug('User prompted: Confirm button reset.')
        _msg = 'You are about to reset button locks!\nDo you wish to continue?'
        _style = 'INC_CANCEL'
        dlg = MyMessageDialog(self, msg=_msg, myStyle=_style)
        dlg.OnSetBorderColour(wx.Colour(255,165,0,255))
        if dlg.ShowModal() == wx.ID_OK:
            if self.cbDayStart.GetValue():
                self.parent.btnStartDay.Enable(True)
                self.parent.txtStartDay.SetValue('')
                log.debug(
                    '\'Start of Day\' button enabled by user.')
            if self.cbLunchStart.GetValue():
                self.parent.btnStartLunch.Enable(True)
                self.parent.txtStartLunch.SetValue('')
                log.debug(
                    '\'Start of Lunch\' button enabled by user.')
            if self.cbLunchEnd.GetValue():
                self.parent.btnEndLunch.Enable(True)
                self.parent.txtEndLunch.SetValue('')
                log.debug(
                    '\'End of Lunch\' button enabled by user.')
            if self.cbDayEnd.GetValue():
                self.parent.btnEndDay.Enable(True)
                self.parent.txtEndDay.SetValue('')
                log.debug(
                    '\'End of Day\' button enabled by user.')
        event.Skip()
        self.EndModal(wx.ID_OK)

    def OnCancel(self, event):
        self.EndModal(wx.ID_CANCEL)


class EnterEmail(CustomDialog):
    def __init__(self, parent, eType, *arg, **kw):
        self.eType = eType
        super(EnterEmail, self).__init__(parent, *arg, **kw)
        self.parent = parent
        self.idEmail = wx.NewId()

        self.OnInitUI()
        self.OnInitLayout()
        self.OnBindEvents()
        self.OnSetBorderColour(wx.Colour(205,133,63,255))

    def OnInitUI(self):
        self._panel = wx.Panel(self)

        self.sb = wx.StaticBox(self._panel, label='')

        self.dEmail = wx.TextCtrl(self._panel, id=self.idEmail,
                value='',
                style=wx.TE_PROCESS_ENTER|wx.SIMPLE_BORDER)

        if self.eType == 'DEST':
            self.sb.SetLabel('Confirm destination email address:')
            #log.debug(
            #    'User is updating Destination email address.')
            self.dEmail.SetValue('enter @ default . email')
        elif self.eType == 'CC':
            #log.debug(
            #    'User is updating CC email address.')
            self.sb.SetLabel('Enter email address to CC:')
        elif self.eType == 'BUG':
            self.sb.SetLabel('Confirm email address for Bug Reports:')
            self.dEmail.SetValue('bug @ report . email')

        self.btnOk = _wgt.CustomPB(self._panel, id=wx.ID_OK, label='OK')
        self.btnCancel = _wgt.CustomPB(self._panel, id=wx.ID_CANCEL, label='Cancel')

    def OnInitLayout(self):
        mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        panelSizer = wx.BoxSizer(wx.VERTICAL)

        sbSizer1 = wx.StaticBoxSizer(self.sb, wx.VERTICAL)
        sbSizer1.Add(self.dEmail, 1, wx.ALL|wx.EXPAND, 5)

        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        btnSizer.Add(self.btnOk, 1, wx.ALL, 5)
        btnSizer.Add(self.btnCancel, 1, wx.ALL, 5)
        sbSizer1.Add(btnSizer, 1, wx.ALL|wx.EXPAND, 5)

        panelSizer.Add(sbSizer1, 1, wx.ALL|wx.EXPAND, 3)
        self._panel.SetSizer(panelSizer)

        mainSizer.Add(self._panel, 1, wx.ALL|wx.EXPAND, 2)
        self.SetSizerAndFit(mainSizer)
        self.SetPosition( (400,300) )

    def OnBindEvents(self):
        self._panel.Bind(wx.EVT_MOTION, self.OnMouse)
        self.dEmail.Bind(wx.EVT_TEXT_ENTER, self.OnOk)

    def GetAddress(self):
        emailAddress = self.dEmail.GetValue()
        if '@' in emailAddress and '.' in emailAddress.split('@')[1]:
            return emailAddress
        else:
            return ''

    def OnOk(self, event):
        self.EndModal(wx.ID_OK)


class ViewEmail(CustomDialog):
    def __init__(self, parent, eType, eMail, *arg, **kw):
        self.eType = eType
        self.eMail = eMail
        super(ViewEmail, self).__init__(parent, *arg, **kw)
        self.OnInitUI()
        self.OnInitLayout()
        self.OnSetBorderColour(wx.Colour(55,95,215,255))

    def OnInitUI(self):
        self._panel = wx.Panel(self)

        self.statBox = wx.StaticBox(self._panel, label='')
        self.statEmail = wx.StaticText(self._panel, label='')
        statFont = wx.Font(14, wx.DECORATIVE, wx.NORMAL, wx.NORMAL)
        self.statEmail.SetFont(statFont)
        if self.eType == 'DEST':
            self.statBox.SetLabel('Destination Address:')
        elif self.eType == 'CC':
            self.statBox.SetLabel('\'Carbon Copy\' Address:')
        elif self.eType == 'BUG':
            self.statBox.SetLabel('\'Bug Report\' Address:')

        emailString = wordwrap(
            str(self.eMail),
            350,wx.ClientDC(self._panel))

        self.statEmail.SetLabel(emailString)

        self.btnOk = _wgt.CustomPB(self._panel, id=wx.ID_OK, label='OK')

    def OnInitLayout(self):
        mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        panelSizer = wx.BoxSizer(wx.VERTICAL)

        sbSizer1 = wx.StaticBoxSizer(self.statBox, wx.VERTICAL)
        sbSizer1.Add(self.statEmail, 1, wx.ALL|wx.CENTER, 3)

        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        btnSizer.Add(self.btnOk, 1, wx.ALL|wx.CENTER, 5)

        sbSizer1.Add(btnSizer, 1, wx.ALL|wx.CENTER, 5)

        panelSizer.Add(sbSizer1, 1, wx.ALL|wx.EXPAND, 5)
        self._panel.SetSizer(panelSizer)

        mainSizer.Add(self._panel, 1, wx.ALL|wx.EXPAND, 3)
        self.SetSizerAndFit(mainSizer)

    def OnOk(self, event):
        if self.IsModal():
            self.EndModal(wx.ID_OK)


class AboutWindow(CustomDialog):
    def __init__(self, parent, ver, *arg, **kw):
        self._ver = ver
        super(AboutWindow, self).__init__(parent, *arg, **kw)
        '''Create and define the About dialog window.'''
        self._idHyperlink = wx.NewId()
        self.OnInitWidgets()
        self.OnInitLayout()
        self.OnBindEvents()
        self.CenterOnParent()

    def OnInitWidgets(self):
        self._panel = wx.Panel(self)

        self.titleText = wx.StaticText(self._panel, -1,
            'Clock Punch v%s' % self._ver)

        titleFont = wx.Font(14, wx.DECORATIVE, wx.NORMAL, wx.BOLD)
        self.titleText.SetFont(titleFont)

        self.webLink = _wgt.CustomHyperlink(self._panel,
            self._idHyperlink,
            label = 'Written using Python',
            url = "http://en.wikipedia.org/wiki/Python_programming_language"
            )

        description = wordwrap(
           "This is the Clock Puncher."
           " \n-The original Time Punch application re-written."
           " \nIt allows you to easily send the various time punch"
           " emails at the click of a button."
           "\n\nPlease do not report any issues with this program to DS IT as"
           " it is not an officially supported application.",
           350, wx.ClientDC(self))

        self.descText = wx.StaticText(self._panel, -1, description)

        self.licenseHeader = wx.StaticText(self._panel, -1, 'Licensing:')
        licenseFont = wx.Font(12, wx.DECORATIVE, wx.NORMAL, wx.BOLD)
        self.licenseHeader.SetFont(licenseFont)

        license = wordwrap('Main program: GPL v3'
            '\nPython: PSF License'
            '\npyWin32: PSF License'
            '\nWMI: MIT License'
            '\nwxPython: wxWidgets License',
            350, wx.ClientDC(self))

        self.licenseText = wx.StaticText(self._panel, -1, license)

        self.devHeader = wx.StaticText(self._panel, -1, 'Developer(s):')
        devFont = wx.Font(12, wx.DECORATIVE, wx.NORMAL, wx.BOLD)
        self.devHeader.SetFont(devFont)

        developers = wordwrap(
            'Program Writer:'
            '\nMichael Stover',
            500, wx.ClientDC(self))

        self.devText = wx.StaticText(self._panel, -1, developers)

        self.btnOk = _wgt.CustomPB(self._panel,
            id=wx.ID_OK, label='OK', size=(50,-1))

    def OnInitLayout(self):
        mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        panelSizer = wx.BoxSizer(wx.VERTICAL)
        titleSizer = wx.BoxSizer(wx.HORIZONTAL)
        titleSizer.Add(self.titleText, 0, wx.ALL|wx.EXPAND|wx.ALIGN_CENTER, 2)

        descriptSizer = wx.BoxSizer(wx.VERTICAL)
        descriptSizer.Add(self.descText, 1, wx.ALL|wx.EXPAND, 2)
        descriptSizer.Add(self.webLink, 0, wx.ALL, 2)

        licenseSizer = wx.BoxSizer(wx.VERTICAL)
        licenseSizer.Add(self.licenseHeader, 0, wx.ALL|wx.EXPAND|wx.CENTER, 2)
        licenseSizer.Add(self.licenseText, 1, wx.ALL|wx.EXPAND, 2)

        developSizer = wx.BoxSizer(wx.VERTICAL)
        developSizer.Add(self.devHeader, 0, wx.ALL|wx.EXPAND|wx.CENTER, 2)
        developSizer.Add(self.devText, 1, wx.ALL|wx.EXPAND, 2)

        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        buttonSizer.Add(self.btnOk, 0, wx.ALL|wx.ALIGN_RIGHT, 2)

        panelSizer.Add(titleSizer, 0, wx.ALL|wx.EXPAND|wx.ALIGN_CENTER, 2)
        panelSizer.Add(descriptSizer, 0, wx.ALL|wx.EXPAND, 2)

        panelSizer.Add(wx.StaticLine(
                            self._panel, -1,
                            style=wx.LI_HORIZONTAL
                            ),
                       1, wx.ALL|wx.EXPAND, 5
                       )

        panelSizer.Add(licenseSizer, 0, wx.ALL|wx.EXPAND, 2)

        panelSizer.Add(wx.StaticLine(
                            self._panel, -1,
                            style=wx.LI_HORIZONTAL
                            ),
                       1, wx.ALL|wx.EXPAND, 5
                       )

        panelSizer.Add(developSizer, 0, wx.ALL|wx.EXPAND, 2)
        panelSizer.Add(buttonSizer, 0, wx.ALL|wx.ALIGN_CENTER, 2)

        self._panel.SetSizer(panelSizer)

        mainSizer.Add(self._panel, 1, wx.ALL|wx.EXPAND, 2)
        self.SetSizerAndFit(mainSizer)

    def OnBindEvents(self):
        self.Bind(wx.EVT_HYPERLINK, self.OnLinkClicked, id=self._idHyperlink)
        self._panel.Bind(wx.EVT_MOTION, self.OnMouse)

    def OnLinkClicked(self, event=None):
        evtObj = event.GetEventObject()
        if isinstance(evtObj, wx.HyperlinkCtrl):
            print evtObj.GetURL()

    def OnOk(self, event):
        if self.IsModal():
            self.EndModal(wx.ID_OK)


class HelpWindow(CustomDialog):
    def __init__(self, parent, *arg, **kw):
        super(HelpWindow, self).__init__(parent, *arg, **kw)
        pass


class BugReport(CustomDialog):
    def __init__(self, parent, *arg, **kw):
        super(BugReport, self).__init__(parent, *arg, **kw)
        self.OnSetBorderColour(wx.Colour(153, 0, 0, 255))

        self.OnInitUI()
        self.OnInitLayout()
        self.OnBindEvents()

    def OnInitUI(self):
        self._panel = wx.Panel(self)
        self._panel.SetBackgroundColour(wx.Colour(230,230,230,255))

        self._titleBox = wx.StaticText(
            self._panel, -1, ' Bug Report Submission',
            style=wx.NO_BORDER
        )
        titleFont = wx.Font(12, wx.DECORATIVE, wx.NORMAL, wx.BOLD)
        self._titleBox.SetFont(titleFont)
        self._titleBox.SetBackgroundColour(wx.WHITE)

        summary = wordwrap(
            'Thank you for taking this opportunity to submit '
            'a bug report for this program.\n\nPlease enter a '
            'brief description into the box below. I will '
            'investigate the bug report as soon as possible.'
            '\n\nUpon clicking Submit the program will also '
            'locate and collect any logs for review. These '
            'will be sent as attachments to the email.'
            '\n\nNo personal information is collected.'
            '\nYou will be able to preview the report before '
            'it is sent.',
            350, wx.ClientDC(self))

        self.bugTips = wx.StaticText(self._panel, -1, summary,
            style=wx.NO_BORDER|wx.TE_CENTER
        )
        self.bugTips.SetBackgroundColour(wx.WHITE)

        self.summaryBox = wx.TextCtrl(self._panel, -1, value='',
            style=wx.TE_MULTILINE|wx.SIMPLE_BORDER)

        self.btnOk = _wgt.CustomPB(self._panel, id=wx.ID_OK, label='OK')
        self.btnCancel = _wgt.CustomPB(self._panel, id=wx.ID_CANCEL, label='Cancel')

    def OnInitLayout(self):
        mainSizer = wx.BoxSizer(wx.HORIZONTAL)
        panelSizer = wx.BoxSizer(wx.VERTICAL)

        titleSizer = wx.BoxSizer(wx.HORIZONTAL)
        titleSizer.Add(self._titleBox, 1, wx.ALL|wx.EXPAND, 3)
        panelSizer.Add(titleSizer, 0, wx.EXPAND|wx.CENTRE, 3)

        tipsSizer = wx.BoxSizer(wx.HORIZONTAL)
        tipsSizer.Add(self.bugTips, 1, wx.EXPAND, 3)
        panelSizer.Add(tipsSizer, 1, wx.ALL|wx.EXPAND, 3)

        summarySizer = wx.BoxSizer(wx.HORIZONTAL)
        summarySizer.Add(self.summaryBox, 1, wx.ALL|wx.EXPAND, 3) 
        panelSizer.Add(summarySizer, 1, wx.ALL|wx.EXPAND, 1)       

        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        buttonSizer.Add(self.btnOk, 1, wx.ALL, 3)
        buttonSizer.Add(self.btnCancel, 1, wx.ALL, 3)
        panelSizer.Add(buttonSizer, 0, wx.EXPAND, 1)

        self._panel.SetSizer(panelSizer)

        mainSizer.Add(self._panel, 1, wx.ALL, 3)
        self.SetSizerAndFit(mainSizer)

    def OnBindEvents(self):
        self._panel.Bind(wx.EVT_MOTION, self.OnMouse)
        self._titleBox.Bind(wx.EVT_MOTION, self.OnMouse)
        self.bugTips.Bind(wx.EVT_MOTION, self.OnMouse)

    def OnGetSummary(self, event=None):
        return self.summaryBox.GetValue()

    def OnClose(self, event):
        self.EndModal(event.GetId())


import logging
log = logging.getLogger('root')

log.debug('Dialogs Module %s Initialized.' % __version__)
