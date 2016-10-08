VERSION 5.00
Begin VB.Form fConvertFromWav 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WAV Encoding"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3015
      Begin VB.Image imgOGGLogo 
         Height          =   240
         Left            =   360
         Picture         =   "fWAV2ACCP.frx":0000
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         Caption         =   "Пожалуйста подождите..."
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "fConvertFromWav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents DOS As DOSOutputs
Attribute DOS.VB_VarHelpID = -1

Function Convert(ByVal iFormat As ENCODE_TO_FORMAT_ENUM, _
                 ByVal sWAVPath As String, _
                 ByVal sDestFile As String, _
                 Optional ByVal lTrackIndex As String, _
                 Optional ByVal lTrackCount As Double = -1) As Long
                 
    On Error Resume Next
    
    Dim Tmp As String
    Dim sCmd As String
    Dim vQuality As String
    
    Select Case iFormat
    Case E2_ACCPLUS
        vQuality = Settings.ACCPlus_Bitrate * 1000
        sCmd = "accplus.dll ""%in%"" ""%out%"" --cbr %quality%"
    Case E2_MP3
        If Settings.MP3_BitrateType = 1 Then
            vQuality = Settings.MP3_VariableBitrate
            sCmd = "lame.dll -f -c -o -p -V %quality% ""%in%"" ""%out%"""
        Else
            vQuality = Settings.MP3_Bitrate
            sCmd = "lame.dll -f -c -o -p -b %quality% ""%in%"" ""%out%"""
        End If
    Case E2_OGG
        vQuality = Settings.OGG_Quality
        sCmd = "oggenc2.dll -q %quality% ""%in%"" -o ""%out%"""
    End Select

    Set DOS = New DOSOutputs
    
    Call Show(vbModeless, fMain)
    fMain.Enabled = False
    DoEvents
    
    Tmp = sCmd
    vQuality = Replace(vQuality, ",", ".")
    Tmp = Replace(Tmp, "%quality%", vQuality)
    Tmp = Replace(Tmp, "%in%", sWAVPath)
    Tmp = Replace(Tmp, "%out%", sDestFile)

    Call DeleteFile(sDestFile)
    Call DOS.ExecuteCommand(RCP(App.Path) & Tmp)
    
    Convert = FileLen(sDestFile)
    
    fMain.Enabled = True
    Call Unload(Me)
    Set fConvertFromWav = Nothing

End Function

