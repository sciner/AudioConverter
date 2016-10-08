VERSION 5.00
Begin VB.UserControl TabStrip 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "TabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Toleranto (c) 2007, icq: 194105899
'http://forum.sources.ru/index.php?showuser=13649

Private Const WC_TABCONTROL = "SysTabControl32"

Event Click()

Private Const WM_SETFONT = &H30

Private Const cNull As Long = &H0
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_COMMAND As Long = &H111
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP = &H202

Private Const WS_TABSTOP  As Long = &H10000
Private Const WS_CHILD  As Long = &H40000000
Private Const WS_EX_NOPARENTNOTIFY  As Long = &H4&
Private Const TCS_FOCUSONBUTTONDOWN As Long = &H1000

Private Const TCIF_TEXT As Long = &H1
Private Const TCM_FIRST As Long = &H1300
Private Const TCM_DELETEALLITEMS As Long = (TCM_FIRST + 9)
Private Const TCM_DELETEITEM As Long = (TCM_FIRST + 8)
Private Const TCM_GETCURSEL As Long = (TCM_FIRST + 11)
Private Const TCM_GETITEMA As Long = (TCM_FIRST + 5)
Private Const TCM_GETITEMCOUNT As Long = (TCM_FIRST + 4)
Private Const TCM_INSERTITEMA As Long = (TCM_FIRST + 7)
Private Const TCM_SETCURSEL As Long = (TCM_FIRST + 12)
Private Const TCM_SETITEMA As Long = (TCM_FIRST + 6)

'TYPE
Private Type TC_ITEM
  Mask          As Long
  dwState       As Long
  dwStateMask   As Long
  pszText       As String
  cchTextMax    As Long
  iImage        As Long
  lParam        As Long
End Type

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private hTabStrip   As Long
Private OldCtl      As Long
Private ASM_Ctl()   As Byte

Dim WithEvents Free As cFree
Attribute Free.VB_VarHelpID = -1

Sub Initialize()
  Call Free.Hook(hTabStrip)
End Sub

Private Sub CreateTabStrip(ByVal hParent As Long, tFont As StdFont, ByVal dwHdc As Long)

  Dim hFont As Long
  Dim dwStyle As Long

  dwStyle = WS_CHILD Or WS_TABSTOP Or TCS_FOCUSONBUTTONDOWN Or &H10000000

  'Call InitCommonControls

  hTabStrip = CreateWindowEx(WS_EX_NOPARENTNOTIFY, WC_TABCONTROL, "TolerantoSoft", dwStyle, 0, 0, ScaleWidth, ScaleHeight, hParent, 0, App.hInstance, ByVal 0)

  If hTabStrip <> 0 Then
    hFont = CreateFont(-tFont.Size * GetDeviceCaps(dwHdc, 90) \ 72, 0, 0, 0, IIf(70, tFont.Bold, 0), tFont.Italic, tFont.Underline, tFont.Strikethrough, 1, 0, 0, 0, 0, tFont.Name)
    Call SendMessage(hTabStrip, WM_SETFONT, hFont, 1&)
    Call DestroyWindow(hFont)
  End If

  Set Free = New cFree
  Call mIOleInPlaceActivate.InitIPAO(Me)

End Sub

Private Sub Move(tsLeft As Long, tsTop As Long, tsWidth As Long, tsHeight As Long)
  Call MoveWindow(hTabStrip, tsLeft, tsTop, tsWidth, tsHeight, 1)
End Sub

Public Sub Add(Optional tsText As String)
  Dim tsItem As TC_ITEM
  tsItem.Mask = TCIF_TEXT
  tsItem.pszText = tsText
  Call SendMessage(hTabStrip, TCM_INSERTITEMA, 10, tsItem)
End Sub

Public Sub Remove(index As Long)
  Call SendMessage(hTabStrip, TCM_DELETEITEM, index, 0)
End Sub

Public Sub RemoveAll()
  Call SendMessage(hTabStrip, TCM_DELETEALLITEMS, 0, 0)
End Sub

Property Get ListCount() As Long
  ListCount = SendMessage(hTabStrip, TCM_GETITEMCOUNT, 0, 0)
End Property

Property Get ListIndex() As Long
  ListIndex = SendMessage(hTabStrip, TCM_GETCURSEL, 0, 0)
End Property

Property Let ListIndex(index As Long)
 Call SendMessage(hTabStrip, TCM_SETCURSEL, index, 0)
End Property

Property Get List(index As Long) As String
Dim tsItem As TC_ITEM
 tsItem.Mask = TCIF_TEXT
 tsItem.pszText = String$(255, Chr(0))
 tsItem.cchTextMax = 255
 Call SendMessage(hTabStrip, TCM_GETITEMA, index, tsItem)
 List = Left$(tsItem.pszText, (InStr(1, tsItem.pszText, Chr$(0)) - 1))
End Property

Property Let List(index As Long, tsText As String)
  Dim tsItem As TC_ITEM
  tsItem.Mask = TCIF_TEXT
  tsItem.pszText = tsText
  Call SendMessage(hTabStrip, TCM_SETITEMA, index, tsItem)
End Property

Public Sub Destroy()
 If OldCtl Then Call SetWindowLong(hTabStrip, &HFFFC, OldCtl)
 Call DestroyWindow(hTabStrip)
End Sub

Private Sub Free_Message(ByVal dwHwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
  Select Case uMsg
  Case WM_MOUSEACTIVATE
    Call pvSetIPAO(Me)
  Case WM_LBUTTONUP
    'zLVCallBack = CallWindowProc(OldCtl, hwnd, uMsg, wParam, lParam)
    RaiseEvent Click
  End Select
  'zLVCallBack = CallWindowProc(OldCtl, hwnd, uMsg, wParam, lParam)
End Sub

Private Sub UserControl_Initialize()
  UserControl.ScaleMode = vbPixels
  Call CreateTabStrip(UserControl.hWnd, UserControl.Font, UserControl.hDC)
End Sub

Private Sub UserControl_Resize()
  Call Move(0, 0, ScaleWidth, ScaleHeight)
End Sub

Private Sub UserControl_Terminate()
  Call Free.Unhook
   Call TerminateIPAO(Me)
End Sub

Property Get hWnd() As Long
  hWnd = hTabStrip
End Property

'------------------
Private Sub UserControl_GotFocus()
  'После создания контрола:
  'Call mIOleInPlaceActivate.InitIPAO(m_uIPAO, Me)
  'Free.Hook(UserControl.hWnd)
  If (Me.hWnd) Then
    Call SetFocusAPI(Me.hWnd)
  End If
End Sub

Friend Function pvEdithWnd() As Long
    If (Me.hWnd) Then pvEdithWnd = hTabStrip 'SendMessageLong(Me.hWnd, LVM_GETEDITCONTROL, 0, 0)
End Function
