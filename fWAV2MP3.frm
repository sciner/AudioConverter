VERSION 5.00
Begin VB.Form fWAV2MP3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WAV -> MP3"
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
         Picture         =   "fWAV2MP3.frx":0000
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
Attribute VB_Name = "fWAV2MP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents DOS As DOSOutputs
Attribute DOS.VB_VarHelpID = -1

Function Convert(ByVal WAV As String, ByVal MP3 As String) As Long

  On Error Resume Next

  Dim Tmp As String

  Set DOS = New DOSOutputs

  Call Show(vbModeless, fMain)
  fMain.Enabled = False
  DoEvents

  If Settings.MP3_BitrateType = 1 Then
    Tmp = "lame.dll -f -c -o -p -V " & CStr(Settings.MP3_VariableBitrate) & " """ & WAV & """ """ & MP3 & """"
  Else
    Tmp = "lame.dll -f -c -o -p -b " & CStr(Settings.MP3_Bitrate) & " """ & WAV & """ """ & MP3 & """"
  End If

  Call DeleteFile(MP3)
  Call DOS.ExecuteCommand(RCP(App.Path) & Tmp)

  Convert = FileLen(MP3)

  fMain.Enabled = True
  Call Unload(Me)
  Set fWAV2MP3 = Nothing

End Function

