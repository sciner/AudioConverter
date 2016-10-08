Attribute VB_Name = "ID3v2"
Option Explicit
Option Base 0

'Created by SCINER: lenar2003@mail.ru
'16/06/2006 12:46
'модуль управления ID3v2 тегами в mp3 файлах
'MP3 file ID3v2 tag manage module

Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Type tid3tagheader
  signature As String * 3
  majorversion As Byte
  minorversion As Byte
  flags As Byte
  Size(3) As Byte
End Type

Dim id As tid3tagheader

Private Function InsertToFile(ByVal Path As String, ByVal Data As String, ByVal lPos As Long) As Boolean
  On Error GoTo normalexit
  Const PORT As Long = 1024000
  Dim lSize As Long
  Dim i As Long
  Dim B() As Byte
  Dim hFile As Long
  Dim lFileLen As Long
  Dim FF As Long
  FF = FreeFile
  Open Path For Binary As #FF
    lFileLen = LOF(FF)
    If lFileLen > 0 Then
      lSize = lFileLen - lPos
      If lSize > 0 Then
        ReDim B(lSize)
        Get #FF, lPos, B
        Put #FF, lPos + Len(Data), B
        Put #FF, lPos, Data
        InsertToFile = True
      End If
    End If
normalexit:
  If FF <> 0 Then Close #FF
  'hFile = CreateFile(Path, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
  'If hFile <> -1 Then
  '  lFileLen = GetFileSize(hFile, ByVal 0&)
  '  If lFileLen > 0 Then
  '    lSize = PORT
  '    For i = lFileLen To lPos Step -PORT
  '      MsgBox i
  '    Next
  '  End If
  'End If
  'If hFile <> -1 Then Call CloseHandle(hFile)
End Function

Private Function CutFromFile(ByVal Path As String, ByVal lStart As Long, ByVal lLength As Long) As Boolean
  Const PORT As Long = 1024000
  Dim hFile As Long
  Dim i As Long
  Dim B() As Byte
  Dim j As Long
  Dim lFileLen As Long
  Dim lSize As Long
  Dim lOperationSize As Long
  Dim lPos As Long
  hFile = CreateFile(Path, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
  If hFile <> -1 Then
    lFileLen = GetFileSize(hFile, ByVal 0&)
    If (lFileLen < 1) Or (lFileLen < lStart) Then
      CutFromFile = True
      GoTo normalexit
    End If
    ReDim B(PORT - 1)
    For i = lStart + lLength - 1 To lFileLen Step PORT
      Call SetFilePointer(hFile, i, 0, 0)
      lSize = PORT
      If i + PORT > lFileLen Then
        lSize = lFileLen - i
        If lSize > 0 Then ReDim B(lSize - 1)
      End If
      If lSize > 0 Then Call ReadFile(hFile, B(0), ByVal lSize, lOperationSize, ByVal 0&)
      If lOperationSize <> lSize Then GoTo normalexit
      Call SetFilePointer(hFile, i - lLength, 0, 0)
      WriteFile hFile, B(0), lSize, lOperationSize, ByVal 0&
      If lOperationSize <> lSize Then GoTo normalexit
    Next
    CutFromFile = SetEndOfFile(hFile) <> 0
  End If
normalexit:
  If hFile <> -1 Then Call CloseHandle(hFile)
End Function

Function RemoveId3v2(ByVal Path As String) As Boolean
  Dim ID3 As String
  ID3 = GetID3v2(Path)
  If ID3 = vbNullString Then
    RemoveId3v2 = True
  Else
    RemoveId3v2 = CutFromFile(Path, 1, Len(ID3))
  End If
End Function

Function SetID3v2(ByVal Path, ByVal Id3v2Tag As String) As Boolean
  If VBA.Left$(Id3v2Tag, 3) = "ID3" Then
    If Not RemoveId3v2(Path) Then Exit Function
    SetID3v2 = InsertToFile(Path, Id3v2Tag, 1)
  End If
End Function

Function GetID3v2(ByVal Path As String) As String
  'for version id3v2.3.x
  Dim FF&
  Dim B() As Byte
  Dim lTagSize As Long
  Dim lTrackSize As Long
  FF = FreeFile
  Open Path For Binary As #FF
    Get #FF, 1, id
    If id.signature <> "ID3" Then GoTo normalexit
    If id.majorversion <> 3 Then GoTo normalexit
    lTagSize = bytetolong(id.Size)
    lTrackSize = LOF(FF) - lTagSize
    If lTagSize < 1 Then GoTo normalexit
    ReDim B(10 + lTagSize - 1)
    Get #FF, 1, B
    GetID3v2 = StrConv(B, vbUnicode)
normalexit:
  Close #FF
End Function

'----------------------------------------------------------------------------------------------------
'   purpose: extract the frame size back to a long
'   require: ubound(bytearray) = 3
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Function bytetolong(ByRef bytearray() As Byte) As Long
  'gs07312001 -   replaced with a loop
  '    bytetolong = bytearray(0) * (2 ^ 21)
  '    bytetolong = bytetolong + bytearray(1) * (2 ^ 14)
  '    bytetolong = bytetolong + bytearray(2) * (2 ^ 7)
  '    bytetolong = bytetolong + bytearray(3) * (2 ^ 0)
  Dim idx As Integer
  bytetolong = 0
  For idx = 0 To 3
    bytetolong = bytetolong + (bytearray(idx) * (2 ^ ((3 - idx) * 7)))
  Next idx
End Function


