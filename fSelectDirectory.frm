VERSION 5.00
Begin VB.Form fSelectDirectory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Выберите директорию с файлами"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "fSelectDirectory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTypes 
      Height          =   2985
      ItemData        =   "fSelectDirectory.frx":000C
      Left            =   4200
      List            =   "fSelectDirectory.frx":000E
      MultiSelect     =   1  'Simple
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.CheckBox chkSubDirs 
      Caption         =   "Поиск во вложенных папках"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отме&на"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Типы:"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4380
      Width           =   1575
   End
End
Attribute VB_Name = "fSelectDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim Types As Collection
Dim WithEvents f As cFileFind
Attribute f.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub InitTypes()
  Dim i As Long
  Set Types = New Collection
  For i = 0 To lstTypes.ListCount - 1
    If lstTypes.Selected(i) Then Call Types.Add(VBA.LCase$(lstTypes.List(i)))
  Next
End Sub

Private Function IsSupportedFile(ByVal Path As String) As Boolean
  Dim i As Long
  Path = VBA.LCase$(Path)
  For i = 1 To Types.Count
    If Path Like Types.Item(i) Then
      IsSupportedFile = True
      Exit Function
    End If
  Next
End Function

Private Sub DestroyTypes()
  Set Types = Nothing
End Sub

Private Sub SaveTypes()
  Dim Tmp As String
  Dim i As Long
  Tmp = ";"
  For i = 0 To lstTypes.ListCount - 1
    If lstTypes.Selected(i) Then Tmp = Tmp & lstTypes.List(i) & ";"
  Next
  SETS("types") = Tmp
End Sub

Private Sub LoadTypes()
  Dim Tmp As String
  Dim i As Long
  Call SupportFileTypes.FillListBox(lstTypes)
  Tmp = ";"
  For i = 0 To lstTypes.ListCount - 1
    Tmp = Tmp & lstTypes.List(i) & ";"
  Next
  Tmp = SETS("types", Tmp)
  For i = 0 To lstTypes.ListCount - 1
    lstTypes.Selected(i) = InStr(Tmp, lstTypes.List(i)) > 0
  Next
End Sub

Private Sub cmdOk_Click()
  Dim i As Long
  Dim sPath As String
  Call InitTypes
  Call SaveTypes
  sPath = RCP(Dir1.Path)
  chkSubDirs.Enabled = False
  Dir1.Enabled = False
  Drive1.Enabled = False
  cmdOk.Enabled = False
  cmdCancel.Enabled = False
  lstTypes.Enabled = False
  lstTypes.BackColor = BackColor
  Enabled = False
  Dir1.BackColor = BackColor
  Drive1.BackColor = BackColor
  SETS("SubFolders") = chkSubDirs
  Set f = New cFileFind
  Call f.FindFiles(sPath, "*.*", , chkSubDirs.value = 1)
  ChDir sPath
  Dim Tmp As String
  For i = 1 To f.Count
    If IsSupportedFile(f.Files(i)) Then
      Call fMain.AddFile(f.Files(i))
      lblProgress.Caption = "Загружено " & CStr(Int(i / f.Count * 100)) & "%"
    End If
    DoEvents
  Next
  Call DestroyTypes
  Set f = Nothing
  Unload Me
End Sub

Private Sub Drive1_Change()
  On Error Resume Next
  Dir1.Path = Drive1
End Sub

Function GetDir() As String
  Show vbModal, fMain
End Function

Private Sub F_Search(ByVal Path As String)
  Static x As Long
  x = x + 1
  If x Mod 10 = 0 Then
    lblProgress.Caption = VBA.Left$(Path, 16) & "..."
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Enabled Then
    Select Case KeyCode
    Case 27
      Unload Me
    Case Else
    End Select
  End If
End Sub

Private Sub Form_Load()
  chkSubDirs = Trunc(val(SETS("SubFolders", 0)), 0, 1)
  Call LoadTypes
  'Call ChangeWindowStyle(cmdOk.hWnd)
  'Call ChangeWindowStyle(cmdCancel.hWnd)
  Call ChangeWindowStyle(Dir1.hWnd)
  Call ChangeWindowStyle(lstTypes.hWnd)
End Sub
