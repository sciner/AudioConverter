Attribute VB_Name = "mPathFunctions"
Option Explicit

'Created by SCINER: lenar2003@mail.ru

Public Enum PathElementType
  Path_Disk = 1
  Path_Extension = 2
  Path_FileName = 3
  Path_WithoutExtension = 4
  Path_Root = 5
  Path_RootName = 6
  Path_FileNameWithoutExtension = 7
End Enum

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Function TempDir() As String
  Dim Tmp As String
  Dim lRet As Long
  Tmp = VBA.Space$(260)
  Call GetTempPath(VBA.Len(Tmp), Tmp)
  lRet = InStr(Tmp, vbNullChar)
  If lRet > 0 Then Tmp = VBA.Left$(Tmp, lRet - 1)
  If VBA.Right$(Tmp, 1) <> "\" Then Tmp = Tmp & "\"
  TempDir = Tmp
End Function

Public Function GetTempFile() As String
  Dim Tmp As String
  Dim lRet As Long
  Tmp = VBA.Space$(260)
  Call GetTempFileName(TempDir, "sml", 0, Tmp)
  lRet = InStr(Tmp, vbNullChar)
  If lRet > 0 Then Tmp = VBA.Left$(Tmp, lRet - 1)
  GetTempFile = Tmp
End Function

Public Function RCP(ByVal Path As String) As String
  RCP = Path & IIf(VBA.Right$(Path, 1) = "\", vbNullString, "\")
End Function

Public Function GetFileName(ByVal fp As String) As String
    Do While InStr(fp, "\") > 0
      fp = Mid(fp, InStr(fp, "\") + 1)
    Loop
    GetFileName = fp
End Function


Function GetPathElement(ByVal sPath$, ByVal PathElement As PathElementType)
  Dim Tmp$
  Dim Tp$()
  Dim PathIsPresent As Boolean
  Dim IsFolder As Boolean
  If VBA.Left(sPath, 4) = "\\?\" Then sPath = VBA.Mid$(sPath, 5)
  sPath = Replace(sPath, "/", "\")
  PathIsPresent = InStr(sPath, "\") > 0
  IsFolder = sPath Like "*\"
  If PathElement <> Path_Disk Then
    If sPath Like "?:\" Then Exit Function
  End If
  Select Case PathElement
  Case Path_Disk '= 1
    If sPath Like "?:\*" Then GetPathElement = VBA.Left$(sPath, 3)
  Case Path_Extension '= 2
    If Not IsFolder Then
      If PathIsPresent Then
        Tmp = Split(sPath, "\")(UBound(Split(sPath, "\")))
      Else
        Tmp = sPath
      End If
      If InStr(Tmp, ".") > 0 Then GetPathElement = Split(Tmp, ".")(UBound(Split(Tmp, ".")))
    End If
  Case Path_FileName, Path_FileNameWithoutExtension '= 3
    If IsFolder Then
      Tmp = Split(sPath, "\")(UBound(Split(sPath, "\")) - 1)
    Else
      Tmp = Split(sPath, "\")(UBound(Split(sPath, "\")))
    End If
    If PathElement = Path_FileNameWithoutExtension Then
      If IsFolder Then Tmp = vbNullString
      If InStr(Tmp, ".") > 0 Then Tmp = VBA.Left$(Tmp, Len(Tmp) - 1 - Len(GetPathElement(Tmp, Path_Extension)))
    End If
    GetPathElement = Tmp
  Case Path_WithoutExtension '= 4
    If Not IsFolder Then
      Tmp = StrReverse(sPath)
      Tmp = Split(sPath, "\")(UBound(Split(sPath, "\")))
      If InStr(Tmp, ".") > 0 Then
        Tmp = Split(Tmp, ".")(UBound(Split(Tmp, ".")))
      Else
        Tmp = vbNullString
      End If
      GetPathElement = VBA.Left$(sPath, Len(sPath) - 1 - Len(Tmp))
    End If
  Case Path_Root '= 5
    If PathIsPresent Then
      If IsFolder Then
        Tp = Split(sPath, "\")
        If UBound(Tp) > 1 Then
          ReDim Preserve Tp(UBound(Tp) - 2)
          GetPathElement = Join(Tp, "\") & "\"
        End If
      Else
        Tp = Split(sPath, "\")
        Tp(UBound(Tp)) = vbNullString
        GetPathElement = Join(Tp, "\")
      End If
    End If
  Case Path_RootName '= 6
    If IsFolder Then
      If UBound(Split(sPath, "\")) > 1 Then
        GetPathElement = Split(sPath, "\")(UBound(Split(sPath, "\")) - 2)
      End If
    Else
      If PathIsPresent Then
        GetPathElement = Split(sPath, "\")(UBound(Split(sPath, "\")) - 1)
      End If
    End If
  Case Else
  End Select
End Function
