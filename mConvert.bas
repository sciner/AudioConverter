Attribute VB_Name = "mConvert"
Option Explicit

Function FileType(ByVal Path As String)
  FileType = VBA.LCase$(GetPathElement(Path, Path_Extension))
End Function

Function Convert(ByVal iFormat As ENCODE_TO_FORMAT_ENUM, ByVal PathFrom As String) As String
 
    Dim sz As Long
    Dim ID3v1 As String
    Dim ID3v2 As String
    Dim WavFile As String
    Dim TempFile As String
    Dim ext As String
    Dim isMp3 As Boolean
    Dim isWav As Boolean
    Dim PathTo As String

    If iFormat = E2_ACCPLUS Then ext = "m4a"
    If iFormat = E2_MP3 Then ext = "mp3"
    If iFormat = E2_OGG Then ext = "ogg"
    If iFormat = E2_WAV Then ext = "wav"

    isMp3 = FileType(PathFrom) = "mp3"
    isWav = FileType(PathFrom) = "wav"
    PathTo = GetConvertToPath(PathFrom, ext)
    TempFile = GetTempFile

    If (isMp3) Then
          ID3v1 = GetID3v1(PathFrom)
          ID3v2 = GetID3v2(PathFrom)
    End If

    If isWav Then
        WavFile = PathFrom
    Else
        Call ConvertToWAV(PathFrom, TempFile)
        WavFile = TempFile
    End If

    sz = fConvertFromWav.Convert(iFormat, WavFile, PathTo, 0, 1)
    If Not isWav Then
        Call DeleteFile(TempFile)
    End If

    If isMp3 And (iFormat = E2_MP3) Then
      Call SetID3v1(PathTo, ID3v1)
      Call SetID3v2(PathTo, ID3v2)
    End If

    Convert = PathTo
    On Error Resume Next
    Call Unload(fConvertFromWav)
    Set fConvertFromWav = Nothing
    
End Function

Function ConvertToWAV(ByVal PathFrom As String, ByVal PathTo As String) As String
  Dim sz As Long
  sz = fMP32WAV.Convert(PathFrom, PathTo, True, 0, 1)   'i, LV.ListCount)
  ConvertToWAV = PathTo
  On Error Resume Next
  Call Unload(fMP32WAV)
  Set fMP32WAV = Nothing
End Function

Function GetConvertToPath(ByVal Path As String, ByVal sNewExtension As String) As String
  'CopyToPathMethod = 0 = в папку с исходным файлом
  'CopyToPathMethod = 1 = в выбранную папку
  'boolReplaceOriginal = True = Заменять оригинальный файл, если CopyToPathMethod = 0
  If Settings.SaveTo_Method = 0 And Settings.Replace_Original_File Then GetConvertToPath = GetPathElement(Path, Path_WithoutExtension) & "." & sNewExtension
  If Settings.SaveTo_Method = 0 And Not Settings.Replace_Original_File Then GetConvertToPath = GetPathElement(Path, Path_WithoutExtension) & "_" & CStr(Settings.MP3_Bitrate) & "." & sNewExtension
  If Settings.SaveTo_Method = 1 Then GetConvertToPath = CopyToPath & GetPathElement(Path, Path_FileNameWithoutExtension) & "." & sNewExtension
End Function
