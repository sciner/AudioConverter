Attribute VB_Name = "mHook"
Option Explicit

Public Const PREV_WND_PROC = "Путь"
Private Const GWL_WNDPROC& = (-4)
Public Const WM_GETMINMAXINFO = &H24
Private Const WM_MOVE = &H3
Public Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_MOUSEFIRST = &H200

Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub Hook(ByVal hWnd As Long, Optional ByVal X_Min As Single = 0, Optional ByVal Y_Min As Single = 0, Optional ByVal X_Max As Single = 0, Optional ByVal Y_Max As Single = 0)
  'added by SCINER 13/01/2006 4:38
  'Если программа запущена из под IDE тогда ничего не субкласируем
  'If InIDE Then Exit Sub
  Dim PrevProc As Long
  PrevProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
  Call SetProp(hWnd, PREV_WND_PROC, PrevProc)
  Call SetProp(hWnd, "X_Min", X_Min)
  Call SetProp(hWnd, "Y_Min", Y_Min)
  Call SetProp(hWnd, "X_Max", X_Max)
  Call SetProp(hWnd, "Y_Max", Y_Max)
End Sub

Public Sub Unhook(ByVal hWnd As Long)
  Dim PrevProc As Long
  PrevProc = GetProp(hWnd, PREV_WND_PROC)
  Call SetWindowLong(hWnd, GWL_WNDPROC, PrevProc)
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim MinMax As MINMAXINFO
    Dim PrevProc As Long
    Dim Tmp As String
    Dim lStyle As Long
    Dim F

    PrevProc = GetProp(hw, PREV_WND_PROC)
    Select Case uMsg
    'Case WM_MOUSEFIRST 'WM_MOUSEACTIVATE 'WM_RBUTTONDOWN
      'fMain.Caption = lParam & "; " & wParam
      'fMain.PopupMenu fMain.mnuMainMenu, vbRightButton
    'Case WM_MOUSEWHEEL
    '  wParam = IIf(wParam < 0, -1, 1)
    '  If FormIsLoad("fExplorer") Then
    '      Call fExplorer.cFileBrowser1.MOUSEWHEEL(wParam)
    '      Exit Function
    '  End If
    '  If FormIsLoad("fNavigator") Then
    '    Call fNavigator.MOUSEWHEEL(wParam)
    '  End If
    'Case WM_MOVE
      'If hw = mdiMain.hWnd And mdiMain.WindowState = vbNormal Then
      '  mdiMain.MeLeft = mdiMain.Left
      '  mdiMain.MeTop = mdiMain.Top
      'End If
    Case WM_GETMINMAXINFO
        WindowProc = CallWindowProc(PrevProc, hw, uMsg, wParam, lParam)
        CopyMemory MinMax, ByVal lParam, Len(MinMax)
        Dim XMIN As Long
        Dim YMIN As Long
        Dim XMAX As Long
        Dim YMAX As Long
        XMIN = GetProp(hw, "X_Min")
        YMIN = GetProp(hw, "Y_Min")
        XMAX = GetProp(hw, "X_Max")
        YMAX = GetProp(hw, "Y_Max")
        If XMIN <> 0 Then MinMax.ptMinTrackSize.X = XMIN
        If YMIN <> 0 Then MinMax.ptMinTrackSize.Y = YMIN
        If XMAX <> 0 Then MinMax.ptMaxTrackSize.X = XMAX
        If YMAX <> 0 Then MinMax.ptMaxTrackSize.Y = YMAX
         CopyMemory ByVal lParam, MinMax, Len(MinMax)
        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Case Else
        WindowProc = CallWindowProc(PrevProc, hw, uMsg, wParam, lParam)
    End Select

End Function



