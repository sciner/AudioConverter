VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Event KeyDown(ByVal KeyCode As Long)
Public Event Message(ByVal dwHwnd As Long, uMsg As Long, wParam As Long, lParam As Long)

Private Const WM_KEYDOWN As Long = &H100
Private Const cNull As Long = &H0

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const AsmMain As String = "558BEC83C4FC8D45FC50FF7514FF7510FF750CFF75086800000000B800000000FFD08B45FCC9C21000"

Private OldCtl As Long
Private ASM_Ctl() As Byte
Dim mHwnd As Long

Public Function zLVCallBack(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case uMsg
  'Case WM_KEYDOWN
  '   RaiseEvent KeyDown(wParam)
  End Select
  RaiseEvent Message(hWnd, uMsg, wParam, lParam)
  zLVCallBack = CallWindowProc(OldCtl, hWnd, uMsg, wParam, lParam)
End Function

Public Function Hook(ByVal hWnd As Long)
  mHwnd = hWnd
  Call StartSubclass(ASM_Ctl, hWnd, OldCtl, cNull)
End Function

'Отключение и удаление всего и вся...
Public Sub Unhook()
  If OldCtl Then Call SetWindowLong(mHwnd, &HFFFC, OldCtl)
End Sub

'Вот он прогресс в VB
'Сабклассинг с пом. ASM (автора не знаю...)
Private Sub StartSubclass(ByRef ASM() As Byte, _
                          ByVal hWnd As Long, _
                          ByRef OldWndProc As Long, _
                          Optional ByVal ProcNumber As Long)
  Dim lng As Long
  Dim tPtr As Long
  lng = Len(AsmMain) \ 2&
  ReDim ASM(cNull To lng - vbNull)
  For lng = cNull To lng - vbNull
    ASM(lng) = Val("&H" & Mid$(AsmMain, (lng) * 2& + vbNull, 2&))
  Next
  Call CopyMemory(tPtr, ByVal ObjPtr(Me), 4&)
  Call CopyMemory(lng, ByVal tPtr + &H1C + (4& * ProcNumber), 4&)
  Call CopyMemory(ASM(23), ObjPtr(Me), 4&)
  Call CopyMemory(ASM(28), lng, 4&)
  OldWndProc = SetWindowLong(hWnd, &HFFFC, VarPtr(ASM(0)))
End Sub

Private Sub Class_Terminate()
  Call Unhook
End Sub
