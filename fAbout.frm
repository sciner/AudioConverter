VERSION 5.00
Begin VB.Form fAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О программе..."
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "fAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   2535
      TabIndex        =   5
      Top             =   375
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write WAV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write WAV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   975
      TabIndex        =   3
      Top             =   375
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   4400
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   120
      X2              =   4400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "sciner@yandex.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2205
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   360
      Picture         =   "fAbout.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Written by SCINER:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Standard Cursor IDs
Private Const IDC_ARROW = 32512&
Private Const IDC_IBEAM = 32513&
Private Const IDC_WAIT = 32514&
Private Const IDC_CROSS = 32515&
Private Const IDC_UPARROW = 32516&
Private Const IDC_SIZE = 32640&
Private Const IDC_ICON = 32641&
Private Const IDC_SIZENWSE = 32642&
Private Const IDC_SIZENESW = 32643&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZEALL = 32646&
Private Const IDC_NO = 32648&
Private Const IDC_HAND = 32649&
Private Const IDC_APPSTARTING = 32650&
Private Const IDC_HELP = 32651&

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Sub SetHandCur(isHand As Boolean)
  Call SetCursor(LoadCursor(0, IIf(isHand, IDC_HAND, IDC_ARROW)))
End Sub

Private Sub cmdOk_Click()
  Call Unload(Me)
End Sub

Private Sub Form_Load()
  'Call ChangeWindowStyle(cmdOk.hWnd)
End Sub

Private Sub lblEmail_Click()
  Call ShellExecute(hWnd, "Open", lblEmail.Caption, vbNullString, vbNullString, vbNull)
End Sub
Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call SetHandCur(True)
End Sub
Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call SetHandCur(True)
End Sub
Private Sub lblEmail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call SetHandCur(True)
End Sub
