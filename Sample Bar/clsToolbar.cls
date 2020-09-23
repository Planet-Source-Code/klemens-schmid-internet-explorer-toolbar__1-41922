VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Klemid's &Search Bar"
'*********************************************************************************************
' Authors:     Eduardo Morcillo (author of the original code)
'              Klemens Schmid (seriously changed the code)
' E-Mail:      klemens.schmid@gmx.de
' Web Page:    http://www.klemid.de
'
' Change log:
' 03/12/2000:  Created by Eduardo Morcillo
' 12/25/2002:  First publication by Klemens
'
' Redistribution:
' You are allowed to reuse this code in your own projects. However, I would highly
' appreciate if you mention the authors as the original source of the code.
'
' Open issues:
' - Tabbing into the toolbar. I found no way to tab into the toolbar from outside,
'   i.e. from an IE pane or from another toolbar. The event, which should is dedicated
'   to this is UIActivateIO. It looks like there is a bug in IE because for toolbars
'   this event never gets fired with a fActivate = True.
' - Accelerators. I found no way to jump into my toolbar's combobox by means of a
'   accelerator like Alt+K w/o installing a keyboard hook
' - Resizing. I found no event which gets fired when the toolbar is resized. I would
'   like to resize my textbox synchronously.
'*********************************************************************************************

Option Explicit

Const MODULE_NAME = "CToolbar"

Implements olelib.IOleWindow
Implements olelib.IDeskBand
Implements olelib.IObjectWithSite
Implements olelib.IInputObject
Implements olelib.IPersist
Implements olelib.IOleControlSite
Implements olelib2.IPersistStream
Implements ISubclass

Private m_oSite As olelib.IUnknown           'Band site object
Private m_oIOS As olelib.IInputObjectSite
Private WithEvents m_oIE As InternetExplorer 'IE Object
Attribute m_oIE.VB_VarHelpID = -1

Private m_lMsgWnd As Long                    'Toolbar parent window
Private m_lToolWnd As Long                   'Toolbar window
Private m_lEdit As Long                      'Textbox window
Private m_lEdit2 As Long                     'Textbox window
Private m_oFont As IFont                     'Textbox font

Private m_Buttons() As TBBUTTON              'array of toolbar buttons

'command IDs for buttons
Const ID_FILE = 1
Const ID_SEARCH = 2

'textbox height
Const nControlHeight = 22

Private Sub AddToolbarButtons()
'create the toolbar buttons
Const PROC_NAME = "AddToolbarButtons"
Log "Entered", MODULE_NAME, PROC_NAME

Const strButtonFile = "&Klemid's Sample Bar"
Const strButtonSearch = "&Search At"

ReDim m_Buttons(0 To 2)

'kind of 'File' menu
m_Buttons(0).idCommand = ID_FILE
m_Buttons(0).iString = StrPtr(strButtonFile)
m_Buttons(0).fsStyle = CTBDropDown Or CTBAutoSize
'm_Buttons(0).fsStyle = TBSTYLE_BUTTON + TBSTYLE_AUTOSIZE
m_Buttons(0).fsState = TBSTATE_ENABLED
m_Buttons(0).iBitmap = 0

'width of the separator controls width of the combobox
m_Buttons(1).fsStyle = BTNS_SEP
m_Buttons(1).iBitmap = 160

'drop down menu providing the web sites
m_Buttons(2).idCommand = ID_SEARCH
m_Buttons(2).iString = StrPtr(strButtonSearch)
m_Buttons(2).fsStyle = CTBDropDown Or CTBAutoSize
m_Buttons(2).fsState = TBSTATE_ENABLED
m_Buttons(2).iBitmap = 1

SendMessage m_lToolWnd, TB_ADDBUTTONSW, 3, m_Buttons(0)

End Sub

Private Function GetButtonIndexFromID(ByVal ID As Integer) As Integer
'compute the button index from the ID
Dim i%

For i = 0 To UBound(m_Buttons)
   If m_Buttons(i).idCommand = ID Then
      GetButtonIndexFromID = i
      Exit Function
   End If
Next
   
End Function

Private Sub AddToolbarIcons()
'set the toolbar's image list

Const PROC_NAME = "AddToolbarIcons"
Log "Entered", MODULE_NAME, PROC_NAME

Dim lImgList As Long
Dim lIcon As Long

'Create a ImageList
lImgList = ImageList_Create(16, 16, ILC_COLOR8 Or ILC_MASK, 1, 1)

#If 1 Then
'you can load icons from the resource
lIcon = LoadIcon(App.hInstance, 101&)
ImageList_ReplaceIcon lImgList, -1, lIcon
lIcon = LoadIcon(App.hInstance, 102&)
ImageList_ReplaceIcon lImgList, -1, lIcon
#Else
'or, you can load icons from the VB ImageList control
ImageList_ReplaceIcon lImgList, -1, frmToolbar.ToolbarImages.ListImages(1).ExtractIcon.Handle
ImageList_ReplaceIcon lImgList, -1, frmToolbar.ToolbarImages.ListImages(2).ExtractIcon.Handle
#End If
'tell the toolbar
SendMessageLong m_lToolWnd, TB_SETIMAGELIST, ByVal 0, lImgList
End Sub

Private Function GetButtonWidth(ByVal ID As Integer)
'returns the width or the toolbar button
Dim udtRect As RECT

On Error Resume Next

SendMessage m_lToolWnd, TB_GETRECT, ID, udtRect
GetButtonWidth = udtRect.Right - udtRect.Left

End Function

Private Function GetButtonRect(ByVal ID As Integer) As RECT
'returns the width or the toolbar button
Dim udtRect As RECT

On Error Resume Next

SendMessage m_lToolWnd, TB_GETRECT, ID, udtRect
GetButtonRect = udtRect

End Function

Private Sub GetDropDownPosition(ByVal ID%, ByRef pRect As RECT)
'returns the position where the drop down menu should appear
Dim udtRect As RECT
Dim rc As Long
    
rc = SendMessage(m_lToolWnd, TB_GETITEMRECT, ID, udtRect)
pRect = udtRect
End Sub

Private Sub ClientRectToScreen(ByVal hwnd As Long, ByRef tR As RECT)
'transform window-related coordinates into screen-related coordinates
Dim tP As POINTAPI

tP.x = tR.Left: tP.y = tR.Top
ClientToScreen hwnd, tP
tR.Left = tP.x: tR.Top = tP.y
tP.x = tR.Right: tP.y = tR.Bottom
ClientToScreen hwnd, tP
tR.Right = tP.x: tR.Bottom = tP.y
End Sub

Private Sub CreateTextBox()
'Create the text box window
'Unforunataly how can't reuse a control from the VB form because the
'control must be owned by IE to get all the accelerator stuff working

m_lEdit = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "Type your search string here", ES_AUTOHSCROLL Or WS_VISIBLE Or WS_CHILD Or WS_TABSTOP, 0, 0, 50, nControlHeight, m_lToolWnd, 1, App.hInstance, ByVal 0&)

' Set the textbox font
Set m_oFont = New StdFont
With m_oFont
   .Size = 8
   .Name = "Trebuchet MS"
End With
SendMessage m_lEdit, WM_SETFONT, m_oFont.hFont, ByVal 0&

End Sub

Private Sub CreateToolbar()
'create a toolbar using API calls
Const PROC_NAME = "CreateToolbar"
Log "Entered", MODULE_NAME, PROC_NAME

Dim tTB As TBBUTTON
Dim udtPos As WINDOWPOS
Dim udtRect As RECT
Dim cx&, cy&, nHeight&

'The window that will receive the toolbar messages is the form
m_lMsgWnd = frmToolbar.hwnd

'Create the toolbar window
m_lToolWnd = CreateToolbarEx(m_lMsgWnd, WS_TABSTOP Or WS_CHILD Or _
                        CCS_NODIVIDER Or CCS_NORESIZE Or _
                        TBSTYLE_LIST Or TBSTYLE_TOOLTIPS Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT, _
                        0, 0, 0, 0, tTB, 0, 16, 16, 16, 16, LenB(tTB))

'create toolbar's child controls
AddToolbarIcons
'Make sure we can have drop-down buttons:
SendMessageLong m_lToolWnd, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS
AddToolbarButtons
'add the combo box
On Error Resume Next
With frmToolbar.cboSearch
   m_lEdit = .hwnd
   'make it a child of the toolbar window
   SetParent .hwnd, m_lToolWnd
   nHeight = .Height / Screen.TwipsPerPixelX
End With
'let combo box start after first button
cx = GetButtonWidth(ID_FILE) + 5
udtRect = GetButtonRect(ID_SEARCH)
cy = udtRect.Left - cx - 5
MoveWindow m_lEdit, cx, 0, cy, nHeight, False

'tell the subclassing mechanism which messages we want to handle
AttachMessage Me, m_lMsgWnd, WM_COMMAND
AttachMessage Me, m_lMsgWnd, WM_NOTIFY

End Sub

Private Sub DestroyToolbar()
'stop subclassing
DetachMessage Me, m_lMsgWnd, WM_COMMAND
DetachMessage Me, m_lMsgWnd, WM_NOTIFY
'remove windows
DestroyWindow m_lEdit
DestroyWindow m_lToolWnd
End Sub

Private Function GetButtonCount() As Long
GetButtonCount = SendMessageLong(m_lToolWnd, TB_BUTTONCOUNT, 0, 0)
End Function

Private Sub PressButton(ByVal iButton As Long, ByVal bState As Boolean)
'simulate button press of toolbar button
Dim tBB As TBBUTTON
Dim fState As Long

If iButton >= 0 And iButton < GetButtonCount() Then
   SendMessage m_lToolWnd, TB_GETBUTTON, iButton, tBB
   fState = Abs(bState)
   SendMessageLong m_lToolWnd, TB_PRESSBUTTON, tBB.idCommand, fState
End If

End Sub

Private Sub IDeskBand_CloseDW(ByVal dwReserved As Long)

' Destroy the windows
DestroyToolbar

End Sub

Private Sub IDeskBand_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDeskBand_GetBandInfo(ByVal dwBandID As Long, ByVal dwViewMode As olelib.GetBandInfo_ViewModes, pdbi As olelib.DESKBANDINFO)
Log "Entered", MODULE_NAME, "IDeskBand_GetBandInfo"

Dim sTitle As String

On Error Resume Next

With pdbi
   
   If (.dwMask And DBIM_MINSIZE) = DBIM_MINSIZE Then
      .ptMinSize.x = 200
      .ptMinSize.y = nControlHeight
   End If
   
   If (.dwMask And DBIM_MAXSIZE) = DBIM_MAXSIZE Then
      .ptMaxSize.y = -1
      .ptMaxSize.x = -1
   End If
   
   If (.dwMask And DBIM_ACTUAL) = DBIM_ACTUAL Then
      .ptActual.x = 200
      .ptActual.y = nControlHeight
   End If
   
   If (.dwMask And DBIM_INTEGRAL) = DBIM_INTEGRAL Then
      .ptIntegral.x = 1
      .ptIntegral.y = 1
   End If
   
#If 0 Then
   'we use no title here
   If (.dwMask And DBIM_TITLE) = DBIM_TITLE Then
      sTitle = "&Klemid's Sample Bar"
      MoveMemory .wszTitle(0), ByVal StrPtr(sTitle), LenB(sTitle) + 2
   End If
#End If
   
   If (.dwMask And DBIM_MODEFLAGS) = DBIM_MODEFLAGS Then
      .dwModeFlags = DBIMF_NORMAL
   End If

End With

End Sub

Private Function IDeskBand_GetWindow() As Long
Log "Entered", MODULE_NAME, "IDeskBand_GetWindow"
'the toolbar window becomes the main window of the band
IDeskBand_GetWindow = m_lToolWnd
End Function

Private Sub IDeskBand_ResizeBorderDW(prcBorder As olelib.RECT, ByVal punkToolbarSite As Long, ByVal fReserved As olelib.BOOL)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDeskBand_ShowDW(ByVal fShow As olelib.BOOL)
Log "Entered", MODULE_NAME, "IDeskBand_ShowDW"

If fShow Then
   ShowWindow m_lToolWnd, SW_SHOWNOACTIVATE
Else
   ShowWindow m_lToolWnd, SW_HIDE
End If

End Sub

Private Sub IInputObject_HasFocusIO()
'check whether the combobox has the the focus currently
Const PROC_NAME = "IInputObject_HasFocusIO"
Log "Entered", MODULE_NAME, PROC_NAME

If frmToolbar.HasFocus Then
   'return S_OK
   If Trace > 1 Then Log "Yes, we have the focus", MODULE_NAME, PROC_NAME
Else
   Err.Raise E_FAIL
End If
End Sub

Private Sub IInputObject_TranslateAcceleratorIO(lpMsg As olelib.MSG)
'handle backspace and other keys before IE does
Const PROC_NAME = "IInputObject_TranslateAcceleratorIO"
If Trace > 1 Then Log "Entered (" & MsgToText(lpMsg.message) & ")", MODULE_NAME, PROC_NAME

Select Case lpMsg.message
   
   Case WM_KEYDOWN
      
      If Trace > 1 Then Log "Key code is " & lpMsg.wParam, MODULE_NAME, PROC_NAME
      
      Select Case lpMsg.wParam
      
         Case vbKeyBack
            'forward message to the input control
            SendMessageLong m_lEdit, WM_CHAR, lpMsg.wParam, lpMsg.lParam
            'we cared!
            Exit Sub
            
         Case vbKeyDelete, vbKeyUp, vbKeyDown
            'forward message to the input control
            SendMessageLong m_lEdit, WM_KEYDOWN, lpMsg.wParam, lpMsg.lParam
            'we cared!
            Exit Sub
            
      End Select
      
   Case WM_SYSKEYDOWN, WM_SYSKEYUP
      'we handle our own toolbar accelerators
      'works only when we have the focus
      If Trace > 1 Then Log "Key code is " & lpMsg.wParam, MODULE_NAME, PROC_NAME
      Select Case lpMsg.wParam
         Case vbKeyK:
            Call DropDown(ID_FILE)
            Exit Sub
         Case vbKeyS:
            Call DropDown(ID_SEARCH)
            Exit Sub
      End Select
      
   Case WM_CHAR
      If Trace > 1 Then Log "Key code is " & lpMsg.wParam, MODULE_NAME, PROC_NAME
      If lpMsg.wParam = vbKeyReturn Then
         Log "User pressed return", MODULE_NAME, PROC_NAME
         Exit Sub
      End If
      
End Select

'default is: we didn't care and leave handling up to IE
Err.Raise E_FAIL

End Sub

Private Sub IInputObject_UIActivateIO(ByVal fActivate As olelib.BOOL, lpMsg As olelib.MSG)
'If fActivate Then SetFocus frmToolbar.cboUrls
Log "Entered (" & fActivate & ")", MODULE_NAME, "IInputObject_UIActivateIO"
If fActivate Then
   'unfortunately this never happens
   SetFocus m_lEdit
Else
   m_oIOS.OnFocusChangeIS Me, False
End If
End Sub

Private Sub DropDown(ByVal ID As Long)
'DropDown button has been clicked -> show popup menu
Dim x&, y&, x2&, y2&
Dim r As RECT
Dim rc As Long
Dim lngFlags&
Dim dummyRect As RECT

lngFlags = TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_VERTICAL ' Or TPM_RETURNCMD Or TPM_NONOTIFY
'Get the position relative to the window
Call GetDropDownPosition(GetButtonIndexFromID(ID), r)
ClientRectToScreen m_lToolWnd, r
x2 = r.Left: y2 = r.Bottom
PressButton ID, True
Select Case ID
   Case ID_FILE
      rc = TrackPopupMenu(GetSubmenuByCaption(frmToolbar, frmToolbar.mnuFile.Caption), lngFlags, x2, y2, 0, frmToolbar.hwnd, dummyRect)
   Case ID_SEARCH
      rc = TrackPopupMenu(GetSubmenuByCaption(frmToolbar, frmToolbar.mnuSearch.Caption), lngFlags, x2, y2, 0, frmToolbar.hwnd, dummyRect)
End Select
'un-press button
PressButton ID, False
End Sub

Public Sub OnFocus(ByVal Focused As Boolean)
Log "Entered (" & Focused & ")", MODULE_NAME, "OnFocus"
On Error Resume Next

'this is absolutely required to handle the accelerators
m_oIOS.OnFocusChangeIS Me, Focused

End Sub

Private Sub IObjectWithSite_GetSite(riid As olelib.UUID, ppvSite As stdole.IUnknown)
Dim lErr As Long

' Get the requested interface
lErr = m_oSite.QueryInterface(riid, ppvSite)

If lErr Then Err.Raise lErr

End Sub

Private Sub IObjectWithSite_SetSite(ByVal pUnkSite As stdole.IUnknown)
Dim oSiteOW As IOleWindow

On Error Resume Next

' Store the new site object
Set m_oSite = pUnkSite
Set m_oIOS = pUnkSite
      
If Not m_oSite Is Nothing Then

   'Get the IE instance
   Set m_oIE = FindIESite(m_oSite)
   'initialize the form providing the controls etc
   Call frmToolbar.Init(m_oIE, Me)
   'Create the toolbar
   CreateToolbar
         
Else
   Set m_oIE = Nothing
End If

End Sub

Private Function IOleControlSite_GetExtendedControl() As Object
'Log "Entered", MODULE_NAME, "IOleControlSite_GetExtendedControl"
Err.Raise E_NOTIMPL
End Function

Private Sub IOleControlSite_LockInPlaceActive(ByVal fLock As olelib.BOOL)
'Log "Entered", MODULE_NAME, "IOleControlSite_LockInPlaceActive"
Err.Raise E_NOTIMPL
End Sub

Private Sub IOleControlSite_OnControlInfoChanged()
Log "Entered", MODULE_NAME, "IOleControlSite_OnControlInfoChanged"
Err.Raise E_NOTIMPL
End Sub

Private Sub IOleControlSite_OnFocus(ByVal fGotFocus As olelib.BOOL)
Log "Entered", MODULE_NAME, "IOleControlSite_OnFocus"
End Sub

Private Sub IOleControlSite_ShowPropertyFrame()
Err.Raise E_NOTIMPL
End Sub

Private Sub IOleControlSite_TransformCoords(pPtlHimetric As olelib.POINT, pPtfContainer As olelib.POINTF, ByVal dwFlags As olelib.XFORMCOORDS)
Err.Raise E_NOTIMPL
End Sub

Private Sub IOleControlSite_TranslateAccelerator(pMsg As olelib.MSG, ByVal grfModifiers As olelib.KEYMODIFIERS)
Log "Entered", MODULE_NAME, "IOleControlSite_TranslateAccelerator"
End Sub

Private Sub IOleWindow_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)
Err.Raise E_NOTIMPL
End Sub

Private Function IOleWindow_GetWindow() As Long
IOleWindow_GetWindow = m_lToolWnd
End Function

Private Sub IPersist_GetClassID(pClassID As olelib.UUID)
'Return the CLSID of this class
CLSIDFromProgID "KlemidSampleBar.CToolbar", pClassID
End Sub

Private Sub IPersistStream_GetClassID(pClassID As olelib.UUID)
IPersist_GetClassID pClassID
End Sub

Private Function IPersistStream_GetSizeMax() As Currency
Err.Raise E_NOTIMPL
End Function

Private Sub IPersistStream_IsDirty()
Err.Raise E_NOTIMPL
End Sub

Private Sub IPersistStream_Load(ByVal pStm As olelib2.IStream)

End Sub

Private Sub IPersistStream_Save(ByVal pStm As olelib2.IStream, ByVal fClearDirty As olelib.BOOL)

End Sub

Private Function MsgToText(ByVal MSG As Long) As String
'Translate message ID to constant name for tracing purposes
Select Case MSG
   Case WM_CHAR:        MsgToText = "WM_CHAR"
   Case WM_COMMAND:     MsgToText = "WM_COMMAND"
   Case WM_KEYDOWN:     MsgToText = "WM_KEYDOWN"
   Case WM_KEYUP:       MsgToText = "WM_KEYUP"
   Case WM_NOTIFY:      MsgToText = "WM_NOTIFY"
   Case WM_SYSKEYDOWN:  MsgToText = "WM_SYSKEYDOWN"
   Case WM_SYSKEYUP:    MsgToText = "WM_SYSKEYUP"
   Case Else:           MsgToText = MSG
End Select
End Function

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
'delegate of the window proc for this object
'receives only message dedicated to the window
Const PROC_NAME = "ISubclass_WindowProc"
Dim hdr As NMHDR

Select Case MSG
  
   Case WM_COMMAND
      Select Case wParam \ &H10000
         'we need to catch the focus notifications this way. Use GotFocus and
         'LostFocus on the form doesn't work because VB is too smart. It doesn't
         'invoke GotFocus when you click into another window and come back to
         'the field you've been before.
         Case EN_KILLFOCUS, CBN_KILLFOCUS
            If Trace > 1 Then Log "EN_KILLFOCUS", MODULE_NAME, PROC_NAME
            OnFocus False
            Handled = True
         Case EN_SETFOCUS, CBN_SETFOCUS
            If Trace > 1 Then Log "EN_SETFOCUS", MODULE_NAME, PROC_NAME
            OnFocus True
            Handled = True
      End Select
    
   Case WM_NOTIFY
      'handle drop down of toolbar buttons
      CopyMemory hdr, ByVal lParam, Len(hdr)
      Select Case hdr.code
      Case TBN_DROPDOWN
         'Get the callback
         If (hwnd = m_lMsgWnd) Then
            'get the command ID
            Dim nmTB As NMTOOLBAR_SHORT
            CopyMemory nmTB, ByVal lParam, Len(nmTB)
            'call the object's methiod
            Call DropDown(nmTB.iItem)
            Handled = True
         End If
      End Select
    
End Select

End Function
