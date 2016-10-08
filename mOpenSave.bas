Attribute VB_Name = "mOpenSave"
Option Explicit

Const OFN_ALLOWMULTISELECT = &H200
Const OFN_EXPLORER = &H80000                         '  new look commdlg
Const OFN_EXTENSIONDIFFERENT = &H400

Private Type OpenFileName
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long
Dim IsSave As Boolean

Public SavedFN As String

Function OpenFileName(ByVal hWnd As Long, FName As String, ByVal Patterns As String) As Long
  Dim lRet As Long
  Dim ofn As OpenFileName
  Const MaxFileBufferLength As Long = 32767
  With ofn
    .lStructSize = Len(ofn)
    .hWndOwner = hWnd
    .hInstance = App.hInstance
    .lpstrFilter = Replace(Patterns, "|", vbNullChar) & vbNullChar & vbNullChar
    .lpstrFile = VBA.Space$(MaxFileBufferLength - 1)
    .nMaxFile = MaxFileBufferLength
    .lpstrFileTitle = VBA.Space$(MaxFileBufferLength)
    .nMaxFileTitle = MaxFileBufferLength
    .lpstrInitialDir = FName
    .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
  End With
  If IsSave Then
    lRet = GetSaveFileName(ofn)
  Else
    lRet = GetOpenFileName(ofn)
  End If
  If (lRet) Then
    FName = ofn.lpstrFile
    If InStr(FName, vbNullChar & vbNullChar) > 0 Then
      FName = VBA.Left$(FName, InStr(FName, vbNullChar & vbNullChar) - 1)
    End If
    OpenFileName = ofn.nFilterIndex
  Else
    OpenFileName = -1
  End If
  IsSave = False
End Function

Function SaveFileName(ByVal hWnd As Long, ByVal FName As String, ByVal Patterns As String)
  IsSave = True
  SaveFileName = OpenFileName(hWnd, FName, Patterns)
End Function

