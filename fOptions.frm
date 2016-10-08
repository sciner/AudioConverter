VERSION 5.00
Begin VB.Form fOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройки"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "fOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmDestFiles 
      Caption         =   "Куда сохранять файлы:"
      Height          =   1695
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   4455
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   4215
         TabIndex        =   8
         Top             =   360
         Width           =   4215
         Begin VB.TextBox txtTo 
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   855
            Width           =   3975
         End
         Begin VB.OptionButton optCopyTo 
            Caption         =   "В папку с исходным файлом"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   2895
         End
         Begin VB.OptionButton optCopyTo 
            Caption         =   "В отдельную папку:"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   2
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox chkReplaceOriginal 
            Caption         =   "Заменять исходный файл ( если * > WAV > * )"
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Width           =   3855
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отме&на"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   2505
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2430
      TabIndex        =   5
      Top             =   2505
      Width           =   1095
   End
   Begin VB.CheckBox chkCB 
      Caption         =   "Следить за буфером обмена"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
End
Attribute VB_Name = "fOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  Settings.Replace_Original_File = chkReplaceOriginal.value = 1
  Settings.Clipboard_Monitoring = chkCB.value = vbChecked
  Settings.SaveTo_Method = IIf(optCopyTo(1).value, 1, 0)
  Settings.Save_Directory = RCP(txtTo.TEXT)
  fMain.tmrCB.Enabled = chkCB.value = 1
  Call Unload(Me)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call ChangeWindowStyle(txtTo.hWnd)
    txtTo.TEXT = CopyToPath
    chkReplaceOriginal.value = IIf(Settings.Replace_Original_File, 1, 0)
    optCopyTo(Settings.SaveTo_Method).value = True
    chkCB.value = IIf(Settings.Clipboard_Monitoring, 1, 0)
    Call optCopyTo_Click(0)
End Sub

Private Sub optCopyTo_Click(index As Integer)
  txtTo.Enabled = optCopyTo(1).value
  txtTo.BackColor = IIf(txtTo.Enabled, vbWhite, vbButtonFace)
  chkReplaceOriginal.Enabled = optCopyTo(0).value
End Sub
