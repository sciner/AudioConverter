VERSION 5.00
Begin VB.Form fSelectFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Выбор папки для конвертируемых треков"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "fSelectFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.OptionButton optExit 
      Caption         =   "Выход из программы"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.OptionButton optDefault 
      Caption         =   "Папка по умолчанию %1%"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   5535
   End
   Begin VB.OptionButton optSelect 
      Caption         =   "Выбрать другую папку"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Папка не найдена, либо не доступна, выберите дальнейшее действие:"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "fSelectFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OptIndex As Long

Private Sub cmdOk_Click()
  Unload Me
End Sub

Function GetOption() As Long
  Beep
  Call Show(vbModal)
  GetOption = OptIndex
End Function

Private Sub Form_Load()
  optDefault.Caption = Replace(optDefault.Caption, "%1%", RCP(App.Path))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If optSelect.value Then OptIndex = 0
  If optDefault.value Then OptIndex = 1
  If optExit.value Then OptIndex = 2
End Sub
