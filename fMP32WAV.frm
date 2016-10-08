VERSION 5.00
Begin VB.Form fMP32WAV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Процесс конвертирования..."
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   Icon            =   "fMP32WAV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTotalProgress 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   840
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   14
      Top             =   2160
      Width           =   4935
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "Отме&на"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtMP3Name 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   840
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   0
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Label lblCurrentTrackProgress 
      AutoSize        =   -1  'True
      Caption         =   "Конвертация текущего трека:"
      Height          =   195
      Left            =   840
      TabIndex        =   15
      Top             =   2520
      Width           =   2280
   End
   Begin VB.Label lblConvertTimeValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label lblConvertTime 
      AutoSize        =   -1  'True
      Caption         =   "Время конвертирования:"
      Height          =   195
      Left            =   840
      TabIndex        =   12
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Line lineShadow1 
      BorderColor     =   &H00FFFFFF&
      X1              =   840
      X2              =   5750
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line lineShadow2 
      BorderColor     =   &H00808080&
      X1              =   840
      X2              =   5750
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lvlAllTracksProgressValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   90
   End
   Begin VB.Label lblElapsedValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label lblElapsed 
      AutoSize        =   -1  'True
      Caption         =   "Общее время конвертирования:"
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   2475
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      Caption         =   "Позиция:"
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   840
      Width           =   705
   End
   Begin VB.Label lblStrFile 
      AutoSize        =   -1  'True
      Caption         =   "Файл:"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Время:"
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   540
   End
   Begin VB.Image imgDiskLogo 
      Height          =   480
      Left            =   120
      Picture         =   "fMP32WAV.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblTimeValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   90
   End
   Begin VB.Label lblPosValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   90
   End
   Begin VB.Label lvlAllTracksProgress 
      AutoSize        =   -1  'True
      Caption         =   "Общий процесс конвертирования:"
      Height          =   195
      Left            =   840
      TabIndex        =   11
      Top             =   1680
      Width           =   2625
   End
End
Attribute VB_Name = "fMP32WAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////
' frmWriteWave.frm - Copyright (c) 2002
'                        JOBnik! [Arthur Aminov, ISRAEL]
'                        e-mail: jobnik2k@hotmail.com
'
' Originally Translated from: - writewav.c - Example of Ian Luck
'
' BASS WAVE writer example: MOD/MPx/OGG -> "BASS.WAV"
'//////////////////////////////////////////////////////////////

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type WAVEHEADER_RIFF  '12 bytes
    RIFF As Long                '"RIFF" = &H46464952
    riffBlockSize As Long       'pos + 44 - 8
    riffBlockType As Long       '"WAVE" = &H45564157
End Type

Private Type WAVEHEADER_data  '8 bytes
   dataBlockType As Long        '"data" = &H61746164
   dataBlockSize As Long        'pos
End Type

Private Type WAVEFORMAT     '24 bytes
    wfBlockType As Long         '"fmt " = &H20746D66
    wfBlockSize As Long
    '--- block size begins from here = 16 bytes
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
End Type

Dim wr As WAVEHEADER_RIFF
Dim wf As WAVEFORMAT
Dim wd As WAVEHEADER_data
Dim LenPos, PicW#
Dim bQuit As Boolean
Dim chan As Long, pos As Long, flags As Long
Dim str As Boolean  'stream OR music
Dim buf() As Byte

Private Function InitBass() As Boolean
  'change and set the current path
  'so it won't ever tell you that bass.dll isn't found
  ChDrive App.Path
  ChDir App.Path
  'check if bass.dll is exists
  If Find(RCP(App.Path) & "bass.dll") <> fFindFile Then
      MsgBox "BASS.DLL не найден", vbCritical, "BASS.DLL"
      Exit Function
  End If
  'Check that BASS 1.7 was loaded
  If BASS_GetStringVersion <> "1.7" Then
      MsgBox "BASS версии 1.7 не обнаружена", vbCritical, "BASS.DLL"
      Exit Function
  End If
  'setup output - "no sound" device, 44100hz, stereo, 16 bits
  If (BASS_Init(-2, 44100, BASS_DEVICE_NOTHREAD, fMain.hWnd) = 0) Then
      MsgBox "Ошибка: Невозможно инициализировать устройство", vbCritical, "Digital output"
      Exit Function
  End If
  InitBass = True
End Function

Function Convert(ByVal MP3 As String, _
                 ByVal WAV As String, _
                 Optional ByVal Speedy As Boolean = False, _
                 Optional ByVal lTrackIndex As String, _
                 Optional ByVal lTrackCount As Double = -1) As Long

  Dim lOnePerec As Double
  Dim z As Long
  Dim T As Long
  Dim sz As Long
  Dim oldSz As Long
  Dim lTotalPerec As Double

  If lTrackCount < 0 Then
    MsgBox "Не указано количество треков в списке.", 16
    Exit Function
  End If
  
  lOnePerec = 100 / lTrackCount

  On Error GoTo error_catcher

  Dim FF As Long
  Dim lWavProgress As Double
  
  If Not InitBass Then
    Call BASS_Free
    Exit Function
  End If

  Call Show(vbModeless, fMain)
  fMain.Enabled = False
  DoEvents

  bQuit = False

  FF = FreeFile
  
  PicW = picProgress.ScaleWidth / 100

  'try streaming the file/url
  chan = BASS_StreamCreateFile(BASSFALSE, MP3, 0, 0, BASS_STREAM_DECODE)
  If chan = 0 Then chan = BASS_StreamCreateURL(MP3, 0, BASS_STREAM_DECODE Or BASS_STREAM_RESTRATE, 0)
  If chan > 0 Then
    pos = BASS_StreamGetLength(chan)
    txtMP3Name.TEXT = GetFileName(MP3)
    str = True
  End If

  'try loading the MOD (with sensitive ramping, and calculate the duration)
  If chan = 0 Then
    chan = BASS_MusicLoad(BASSFALSE, MP3, 0, 0, BASS_MUSIC_DECODE Or BASS_MUSIC_RAMPS Or BASS_MUSIC_CALCLEN)
    If chan = 0 Then
      'not a MOD either
      MsgBox "Ошибка: Невозможно открыть файл", vbExclamation
      Exit Function
    Else
      'для трекерной мызыки сохраняем имя трека,
      'для записи в теги результирующего mp3-файла
      modBass.MusicName = BASS_MusicGetNameString(chan)
      txtMP3Name.TEXT = "MOD музыка \" & MusicName & "\ [" & BASS_MusicGetLength(chan, BASSFALSE) & " orders]"
      pos = BASS_MusicGetLength(chan, BASSTRUE)
      str = False
    End If
  End If

  'display the time length
  If (pos) Then
    pos = CLng(BASS_ChannelBytes2Seconds(chan, pos))
    LenPos = pos
    lblTimeValue.Caption = TimeSerial(0, 0, pos)
  Else 'no time length available
    lblPos.Caption = vbNullString
  End If
  
  'конвертирование
  Call picTotalProgress.Cls
  Call picProgress.Cls
  
  DoEvents

  'Set WAV Format
  flags = BASS_ChannelGetFlags(chan)
  wf.wFormatTag = 1
  wf.nChannels = IIf(flags And BASS_SAMPLE_MONO, 1, 2)
  Call BASS_ChannelGetAttributes(chan, wf.nSamplesPerSec, -1, -1)
  wf.wBitsPerSample = IIf(flags And BASS_SAMPLE_8BITS, 8, 16)
  wf.nBlockAlign = wf.nChannels * wf.wBitsPerSample / 8
  wf.nAvgBytesPerSec = wf.nSamplesPerSec * wf.nBlockAlign
  wf.wfBlockType = &H20746D66        '"fmt "
  wf.wfBlockSize = 16
  
  'Set WAV "RIFF" header
  wr.RIFF = &H46464952             '"RIFF"
  wr.riffBlockSize = 0      'after convertion
  wr.riffBlockType = &H45564157    '"WAVE"

  'set WAV "data" header
  wd.dataBlockType = &H61746164     '"data"
  wd.dataBlockSize = 0       'after convertion
  
  'create a file BASS.WAV
  Open WAV For Binary Lock Read Write As #FF
  
    'Write WAV Header to file
    Put #FF, , wr    'RIFF
    Put #FF, , wf    'Format
    Put #FF, , wd    'data
    
    pos = 0
    ReDim buf(199999) As Byte

    T = GetTickCount
    While BASS_ChannelIsActive(chan)
      sz = BASS_ChannelGetData(chan, buf(0), 200000) - 1
      If (sz <> oldSz) Then ReDim Preserve buf(sz) As Byte
      oldSz = sz
      'write data to WAV file
      Put #FF, , buf
      pos = BASS_ChannelGetPosition(chan)
      If z Xor 100 = 0 Then
        lWavProgress = (BASS_ChannelBytes2Seconds(chan, pos) / LenPos) * 100
        lWavProgress = ((lOnePerec / 100) * lWavProgress) * 100
        If str Then
          lTotalPerec = ((lTrackIndex / lTrackCount * 100) + lWavProgress) / 100
          lvlAllTracksProgressValue.Caption = VBA.Format$(lTotalPerec, "0.00") & " %"
          lblPosValue.Caption = TimeSerial(0, 0, BASS_ChannelBytes2Seconds(chan, pos))
          picTotalProgress.Line (0, 0)-Step((picTotalProgress.ScaleWidth / 100) * (lTotalPerec + (lTrackIndex / lTrackCount * 100)), picTotalProgress.ScaleHeight), vbHighlight, BF
          picProgress.Line (0, 0)-Step(PicW * ((BASS_ChannelBytes2Seconds(chan, pos) / LenPos) * 100), 30), vbHighlight, BF
        Else
          lblPosValue.Caption = GetLoWord(pos) & ":" & GetHiWord(pos)
        End If
        lblConvertTimeValue = Format$((GetTickCount - T) / 1000, "0.00") & " сек"
        lblElapsedValue = Format$((GetTickCount - GlobalTimer) / 1000, "0.00") & " сек"
      End If
      DoEvents  'in case you want to exit...
      If Not Speedy Then If z Xor 25 Then Call Sleep(1) 'don't hog the CPU too much :)
      z = z + 1
      If bQuit Then GoTo exit_and_delete_wav
    Wend

    'complete WAV header
    wr.riffBlockSize = pos + 44 - 8
    wd.dataBlockSize = pos
    
    On Error Resume Next
    Put #FF, 5, wr.riffBlockSize
    Put #FF, 41, wd.dataBlockSize

exit_and_delete_wav:
  Close #FF
  FF = 0
  If bQuit Then DeleteFile WAV
  
  On Error Resume Next
  Call Err.Clear
  Convert = FileLen(WAV)
  If Err.Number <> 0 Then Convert = 0
  
  str = False

unload_all:
  If (FF <> 0) Then Close #FF
  Call BASS_Free
  fMain.Enabled = True
  Call Unload(Me)
  Set fMP32WAV = Nothing
  
  Exit Function
error_catcher:
  Call MsgBox(Err.Description, vbCritical, App.Title)
  GoTo unload_all

End Function

Private Sub cmdOk_Click()
  bQuit = True
  GlobalCancel = True
End Sub

Private Sub Form_Load()
  Call ChangeWindowStyle(picProgress.hWnd)
  Call ChangeWindowStyle(picTotalProgress.hWnd)
  Call ChangeWindowStyle(cmdOk.hWnd)
End Sub
