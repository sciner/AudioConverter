Attribute VB_Name = "mListView"
Option Explicit
Option Compare Text

'subclass---
Private Const WM_KEYDOWN = &H100
Private Const WM_NCHITTEST = &H84
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_NCRBUTTONUP = &HA5
Private Const LVM_FIRST                    As Long = &H1000
Private Const LVM_HITTEST                  As Long = (LVM_FIRST + 18)
Private Const GWL_WNDPROC& = (-4)
Private Const WM_MOVE = &H3
Private Const WM_MOUSEFIRST = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Type LVHITTESTINFO
  pt       As POINTAPI
  flags    As Long
  iItem    As Long
  iSubItem As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Created by SCINER: lenar2003@mail.ru

Private Const LVIF_TEXT           As Long = &H1
Private Const LVIF_IMAGE          As Long = &H2
Private Const LVIF_PARAM          As Long = &H4
Private Const LVIF_STATE          As Long = &H8
Private Const LVIF_INDENT         As Long = &H10
Private Const LVIF_GROUPID        As Long = &H100
Private Const LVIF_COLUMNS        As Long = &H200
'Private Const LVM_FIRST           As Long = &H1000
Private Const LVM_GETITEMTEXT     As Long = (LVM_FIRST + 45)

Private Type LVFINDINFO
  flags       As Long
  psz         As String
  lParam      As Long
  pt          As POINTAPI
  vkDirection As Long
End Type

Private Type LVITEM_lp
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    stateMask  As Long
    pszText    As Long
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public m_PRECEDE As Long
Public m_FOLLOW As Long
Public m_lColumn As Long
Public lSortAs As Long
Private m_uLVFI As LVFINDINFO
Private m_uLVI As LVITEM_lp

Private Function pvGetItemText(ByVal hWnd As Long, ByVal lParam As Long) As String
  Dim lIdx   As Long
  Dim a(261) As Byte
  Dim lLen   As Long
  With m_uLVI
      .mask = LVIF_TEXT
      .pszText = VarPtr(a(0))
      .cchTextMax = UBound(a)
      .iSubItem = m_lColumn
  End With
  lLen = SendMessage(hWnd, LVM_GETITEMTEXT, lParam, m_uLVI)
  pvGetItemText = Left$(StrConv(a(), vbUnicode), lLen)
End Function

Function hSortFunc(ByVal lParam1 As Long, _
                   ByVal lParam2 As Long, _
                   ByVal hWnd As Long) As Long

  Select Case lSortAs
  Case stText: hSortFunc = IIf(pvGetItemText(hWnd, lParam1) > pvGetItemText(hWnd, lParam2), m_PRECEDE, m_FOLLOW)
  Case stIndex: hSortFunc = IIf(lParam1 > lParam2, m_PRECEDE, m_FOLLOW)
  Case stDate: hSortFunc = IIf(CDate(pvGetItemText(hWnd, lParam1)) > CDate(pvGetItemText(hWnd, lParam2)), m_PRECEDE, m_FOLLOW)
  Case stNumber: hSortFunc = IIf(val(pvGetItemText(hWnd, lParam1)) > val(pvGetItemText(hWnd, lParam2)), m_PRECEDE, m_FOLLOW)
  End Select

End Function

'subclass---
Public Sub HookListview(ByVal hWnd As Long)
  Dim PrevProc As Long
  PrevProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
  Call SetProp(hWnd, PREV_WND_PROC, PrevProc)
End Sub

Public Sub UnhookkListview(ByVal hWnd As Long)
  Dim PrevProc As Long
  PrevProc = GetProp(hWnd, PREV_WND_PROC)
  Call SetWindowLong(hWnd, GWL_WNDPROC, PrevProc)
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim PrevProc As Long
    PrevProc = GetProp(hw, PREV_WND_PROC)
    Select Case uMsg
    Case WM_KEYDOWN
      If wParam = 46 Then
        '[del]
        Call fMain.mnuOperations_Click(0)
      End If
    'Case WM_NCHITTEST 'WM_RBUTTONUP, WM_RBUTTONDBLCLK
    '  Dim lItm As Long
    '  lItm = pvItemHitTest(hw)
    '  fMain.Caption = lItm & "\" & CStr(Rnd) & "\" & CStr(wParam) & "\" & CStr(lParam)
    Case Else
        WindowProc = CallWindowProc(PrevProc, hw, uMsg, wParam, lParam)
    End Select
End Function

Private Sub pvUCCoords(ByVal lHwnd As Long, uPoint As POINTAPI)
  Call GetCursorPos(uPoint)
  Call ScreenToClient(lHwnd, uPoint)
End Sub

Private Function pvItemHitTest(ByVal lHwnd As Long) As Long
  Dim uLVHI As LVHITTESTINFO
  Call pvUCCoords(lHwnd, uLVHI.pt)
  pvItemHitTest = SendMessage(lHwnd, LVM_HITTEST, 0, uLVHI)
End Function






