
# _widgets.py

import wx
import wx.lib.platebtn as PB
import wx.lib.agw.flatmenu as FM

__version__ = '1.0.0'

class CustomTextCtrl(wx.TextCtrl):
    def __init__(self, parent, *arg, **kw):
        super(CustomTextCtrl, self).__init__(parent, *arg, **kw)
        self.SetEditable(False)
        self.SetBackgroundColour(wx.WHITE)
        self.SetForegroundColour(wx.BLACK)
        self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        self.Bind(wx.EVT_MOUSE_EVENTS, self.DontTouchMe)
        self.Bind(wx.EVT_PAINT, self.DontTouchMe)

    def DontTouchMe(self, event):
        self.HideNativeCaret()
        event.Skip()


class CustomPB(PB.PlateButton):
    def __init__(self, parent, *arg, **kw):
        super(CustomPB, self).__init__(parent, *arg, **kw)
        self.SetBackgroundColour(wx.WHITE)


class CustomHyperlink(wx.HyperlinkCtrl):
    def __init__(self, parent, *arg, **kw):
        super(CustomHyperlink, self).__init__(parent, *arg, **kw)
        self.SetNormalColour(wx.BLUE)
        self.SetHoverColour(wx.RED)
        self.SetVisitedColour(wx.BLUE)


import logging
log = logging.getLogger('root')

log.debug('Widgets Module Initialized.')
