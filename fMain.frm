VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Write Wave 2.0"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11190
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8085
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPanel 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   11190
      TabIndex        =   0
      Top             =   6390
      Width           =   11190
      Begin VB.PictureBox picFormats 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   2
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   10695
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   10695
         Begin VB.ComboBox cmbVBpp 
            Height          =   315
            ItemData        =   "fMain.frx":0E42
            Left            =   1800
            List            =   "fMain.frx":0E64
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cmbBpp 
            Height          =   315
            ItemData        =   "fMain.frx":0EF3
            Left            =   1800
            List            =   "fMain.frx":0F21
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optBpp 
            Caption         =   "Фиксированный"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton optBpp 
            Caption         =   "Переменный"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.PictureBox picFormats 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   0
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   10695
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   10695
         Begin VB.OptionButton optBppACCPlus 
            Caption         =   "Фиксированный"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cmbACCPlusBr 
            Height          =   315
            ItemData        =   "fMain.frx":0F64
            Left            =   1680
            List            =   "fMain.frx":0F8F
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Качество ACCPlus:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   1410
         End
      End
      Begin VB.PictureBox picFormats 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   1
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   10695
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   10695
         Begin MediaConverter.cSlideBar sldOGGQuality 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   150
            Width           =   3495
            _ExtentX        =   5953
            _ExtentY        =   450
         End
         Begin VB.Label lblQGGQualityLabel 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "5"
            Height          =   255
            Left            =   3840
            TabIndex        =   14
            Top             =   75
            Width           =   615
         End
      End
      Begin VB.PictureBox picFormats 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   3
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   10695
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   10695
      End
      Begin MediaConverter.TabStrip TabStrip1 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   2778
      End
   End
   Begin VB.Timer tmrCB 
      Interval        =   256
      Left            =   240
      Top             =   120
   End
   Begin VB.Menu mnuMainMenu 
      Caption         =   "Файл"
      Begin VB.Menu mnuAdd 
         Caption         =   "Добавить файл..."
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Добавить папку..."
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuWinamp 
         Caption         =   "Винамп"
         Begin VB.Menu mnuAdd0 
            Caption         =   "Добавить все из текущего плейлиста Winamp"
            Enabled         =   0   'False
            Shortcut        =   ^W
         End
         Begin VB.Menu mnuAdd1 
            Caption         =   "Добавить текущий трек из Winamp'а"
            Enabled         =   0   'False
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Очистить список"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvertAll 
         Caption         =   "Конвертировать"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Настройки..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Выход"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Правка"
      Begin VB.Menu mnuOperations 
         Caption         =   "Убрать из списка"
         Index           =   0
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Помощь"
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public LV As cListView
Dim ImgList As New cImgList
Dim OldCB As String

Function FindTrack(ByVal Path As String) As Boolean
  Dim i As Long
  For i = 0 To LV.ListCount - 1
    If LV.SubItemText(i, 0) = Path Then
      FindTrack = True
      Exit Function
    End If
  Next
End Function

Public Sub AddFile(ByVal FileName As String)
  Dim j As Long
  If Len(VBA.Trim$(FileName)) = 0 Then Exit Sub
  If FindTrack(FileName) Then Exit Sub
  j = ImgList.AddIconFromFile(FileName)
  If j = -1 Then Exit Sub
  j = LV.ItemAdd(0, FileName, 0, j)
  LV.SubItemText(j, 1) = FileLen(FileName) \ 1024 & " Кб"
  LV.SubItemText(j, 2) = "загружен"
End Sub

Private Sub cmbACCPlusBr_Click()
    Settings.ACCPlus_Bitrate = val(cmbACCPlusBr.TEXT)
End Sub

Private Sub cmbBpp_Click()
    Settings.MP3_Bitrate = val(cmbBpp.TEXT)
End Sub

Private Sub cmbVBpp_Click()
    Settings.MP3_VariableBitrate = val(cmbVBpp.TEXT)
End Sub

Private Sub Form_Initialize()
  CONFIG_FILE = RCP(App.Path) & App.EXEName & ".ini"
  With SupportFileTypes
    .Add "mp3"
    .Add "wav"
    .Add "mo3"
    .Add "it"
    .Add "mod"
    .Add "xm"
    .Add "s3m"
  End With
End Sub

Private Sub Form_Load()

    Dim i As Long

    Call ShiftMenu(hWnd, 2, "Помощь")
    Set LV = New cListView
    Call LV.Init(hWnd)
    Set LV.ImageList(LVSIL_SMALL) = ImgList
    Call LV.ColumnAdd("Путь", 353)
    Call LV.ColumnAdd("Размер", 90, LVCFMT_RIGHT)
    Call LV.ColumnAdd("Статус", 90, LVCFMT_RIGHT)
    LV.View = LVS_REPORT
    LV.FullRowSelect = True
    LV.FlatScrollBar = True
    Call RestoreList
    Call Hook(hWnd, 320 * 1.5, 240 * 1.5)
    Call BackUpList

    LV.HideSelection = False

    Call TabStrip1.Initialize
    Call TabStrip1.Add("ACCPlus")
    Call TabStrip1.Add("Ogg")
    Call TabStrip1.Add("Mp3")
    Call TabStrip1.Add("Wav")
    TabStrip1.ListIndex = Settings.ConvertToFormat
    Call TabStrip1_Click
    
    Call LoadSettings

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Effect <> 7 Then Exit Sub
  Dim i&
  For i = 1 To Data.Files.Count
    Call AddFile(Data.Files(i))
  Next
  Call BackUpList
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
  Effect = IIf(Effect = 7, Effect, 0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call Settings.SaveSettings

  Call Unhook(hWnd)
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  ScaleMode = vbPixels
  If Me.WindowState = vbMinimized Then Exit Sub
  Call LV.Move(5, 5, ScaleWidth - 10, ScaleHeight - picPanel.Height - 10)
End Sub

Private Sub mnuAbout_Click()
  Call fAbout.Show(vbModal, Me)
  'MsgBox App.FileDescription, 32, App.Title
End Sub

Private Sub mnuAdd_Click(index As Integer)
  Dim lRet As Long
  Dim Tmp As String
  Dim Path As String
  Select Case index
  Case 0 'add file
    lRet = mOpenSave.OpenFileName(hWnd, Tmp, SupportFileTypes.GetForShowDialog & "|Все|*.*")
    If (lRet <> -1) Then
      lRet = InStr(Tmp, vbNullChar)
      If lRet > 0 Then Path = RCP(VBA.Left$(Tmp, lRet - 1))
      Tmp = VBA.Mid$(Tmp, InStr(Tmp, vbNullChar) + 1)
      Do While InStr(Tmp, vbNullChar) > 0
        lRet = InStr(Tmp, vbNullChar)
        Call AddFile(Path & VBA.Left$(Tmp, lRet - 1))
        Tmp = VBA.Mid$(Tmp, lRet + 1)
      Loop
      Call AddFile(Path & Tmp)
    End If
  Case 1 'add folder
    Call fSelectDirectory.GetDir
    Set fSelectDirectory = Nothing
  End Select
  Call BackUpList
End Sub

Private Sub mnuClear_Click()
  Call LV.Clear
  Call BackUpList
End Sub

Function FileIsEnableForRead(ByVal lFilePath$)
  On Error Resume Next
  Dim FF&
  Dim B() As Byte
  ReDim B(0 To 0)
  Call Err.Clear
  FF = FreeFile
  Open lFilePath For Binary As #FF
  Get #FF, 1, B
  FileIsEnableForRead = Err.Number = 0
  If FileIsEnableForRead Then Close #FF
End Function

Function FileIsEnable(ByVal lFilePath$)
  On Error Resume Next
  Dim FF&
  Call Err.Clear
  FF = FreeFile
  Open lFilePath For Binary Lock Read Write As #FF
  FileIsEnable = Err.Number = 0
  If FileIsEnable Then Close #FF
End Function

Function MP3_WAV(ByVal FileName As String) As String
  Dim WavFile$
  If FileType(FileName$) <> "wav" Then WavFile$ = ConvertToWAV(FileName$, GetConvertToPath(FileName$, "wav"))
  MP3_WAV = WavFile$
End Function

Private Sub mnuConvert_Click(index As Integer)

''    fMP32WAV.Convert(MP3Src, WAVto, True, i, LV.ListCount)
''    sz = fWAV2MP3.Convert(WAVto, TempTo, TAG)
''    TAG = ReadTAG(MP3Src)
''    Call BackUpList
'
'  Exit Sub
'
'  Dim sz As Long
'  Dim Tmp As String
'  Dim TAG As String
'  Dim MP3Src As String
'  Dim MP3to As String
'  Dim WAVto As String
'  Dim TempTo As String
'  Dim FF As Long
'
'  GlobalCancel = False
'  GlobalTimer = GetTickCount
'  TempTo = GetTempFile
'  FF = FreeFile
'  Open TempTo For Binary As #FF
'  Close #FF
'
'  For i = 0 To LV.ListCount - 1
'
'    MP3Src = LV.SubItemText(i, 0)
'
'    If Not FileIsEnable(MP3Src) And boolReplaceOriginal Then
'      Call MsgBox("Файл не доступен для блокировки" & vbCrLf & MP3Src & vbCrLf & Err.Description, 16)
'      GoTo NEXT_FILE
'    ElseIf Not FileIsEnableForRead(MP3Src) Then
'      Call MsgBox("Файл не доступен для чтения" & vbCrLf & MP3Src & vbCrLf & Err.Description, 16)
'      GoTo NEXT_FILE
'    End If
'
'    'If DeleteFile(TempTo) = 0 Then
'    '  Call MsgBox("Файл занят." & vbCrLf & TempTo, 16)
'    '  GoTo NEXT_FILE
'    'End If
'    Let LV.SubItemText(i, 2) = "копируется..."
'    DoEvents
'
'    If CopyFile(MP3Src, TempTo, False) = 0 Then
'      Call MsgBox("Не удалось скопировать файл." & vbCrLf & MP3Src & vbCrLf & TempTo, 16)
'      GoTo NEXT_FILE
'    End If
'
'    '-> MP3to
'    '-> WAVto
'    '-> MP3Src
'    '-> TempTo
'
'    Let LV.SubItemText(i, 2) = "конвертируется..."
'    DoEvents
'
'    Select Case CopyToPathMethod
'    Case 0 'в папку с исходным файлом
'      If boolReplaceOriginal Then
'        MP3to = MP3Src
'      Else
'        MP3to = GetPathElement(MP3Src, Path_WithoutExtension) & "_" & CStr(GetBpp) & ".mp3"
'      End If
'      WAVto = GetPathElement(MP3Src, Path_WithoutExtension) & ".wav"
'      If WAVto = MP3Src Then
'        WAVto = GetPathElement(MP3Src, Path_WithoutExtension) & "_2.wav"
'      End If
'    Case Else 'в особую папку
'      If GetPathElement(MP3Src, Path_Extension) = "mp3" Then
'        MP3to = CopyToPath & GetPathElement(MP3Src, Path_FileName)
'      Else
'        MP3to = CopyToPath & GetPathElement(MP3Src, Path_FileNameWithoutExtension) & ".mp3"
'      End If
'      WAVto = CopyToPath & GetPathElement(MP3Src, Path_FileNameWithoutExtension) & ".wav"
'    End Select
'    DoEvents
'
'    'MsgBox MP3to & vbCrLf & _
'           WAVto & vbCrLf & _
'           MP3Src & vbCrLf & _
'           TempTo, 48
'    'GoTo NEXT_FILE
'
'    Select Case index
'    Case 0
'      '+ok
'      Call DeleteFile(WAVto)
'      sz = fMP32WAV.Convert(MP3Src, WAVto, True, i, LV.ListCount)
'    Case 1
'      If GetPathElement(MP3Src, Path_Extension) = "wav" Then
'        WAVto = MP3Src
'        sz = FileLen(MP3Src)
'      Else
'        WAVto = GetTempFile
'        MusicName = vbNullString
'        sz = fMP32WAV.Convert(MP3Src, WAVto, True, i, LV.ListCount)
'        If MusicName = vbNullString Then
'          TAG = ReadTAG(MP3Src)
'        Else
'          TAG = VBA.String$(128, vbNullChar)
'          Mid(TAG, 1, 3) = "TAG"
'          Mid(TAG, 4, 30) = VBA.Left$(MusicName, 30)
'          Mid(TAG, 94, 4) = Year(Now)
'          Mid(TAG, 98, 17) = "lenar2003@mail.ru"
'        End If
'      End If
'      If (sz > 0) Then
'        Call DeleteFile(TempTo)
'        If FormIsLoad("fWAV2MP3") Then
'          Call Unload(fWAV2MP3)
'          Set fWAV2MP3 = Nothing
'        End If
'        sz = fWAV2MP3.Convert(WAVto, TempTo, TAG)
'        Call FileCopy(TempTo, MP3to)
'      End If
'      Call DeleteFile(WAVto)
'      Call DeleteFile(TempTo)
'    End Select
'    If sz > 0 Then
'      LV.SubItemText(i, 2) = sz \ 1024 & " Кб - сконвертирован"
'    Else
'      LV.SubItemText(i, 2) = "ошибка"
'    End If
'    Call BackUpList
'
'NEXT_FILE:
'    DoEvents
'    If GlobalCancel Then Exit For
'
'  Next
'
'  Set fMP32WAV = Nothing
'  Set fWAV2MP3 = Nothing
'
'  Call Beep

End Sub

Private Sub mnuConvertAll_Click()

  Dim i As Long
  Dim TAG As String
  Dim FileName As String
  Dim ConvertedFile As String
  Dim index As Long
  
  index = Me.TabStrip1.ListIndex

  'initizlize timer
   GlobalCancel = False
   GlobalTimer = GetTickCount

  'enumerate list and convert all items
  For i = 0 To LV.ListCount - 1
    'get next item from list
    FileName = LV.SubItemText(i, 0)
    'select convert algorithm
    Select Case index
    Case 0: ConvertedFile = Convert(E2_ACCPLUS, FileName)
    Case 1: ConvertedFile = Convert(E2_OGG, FileName)
    Case 2: ConvertedFile = Convert(E2_MP3, FileName)
    Case 3: ConvertedFile = MP3_WAV(FileName)
    Case Else
    End Select
    'if user press cancel on any enable cancel button
    If GlobalCancel Then Exit For
  Next

  Beep

End Sub

Private Sub mnuEdit_Click()
  Dim i&
  For i = 0 To mnuOperations.Count - 1
    mnuOperations(i).Enabled = LV.ListCount > 0
  Next
End Sub

Private Sub mnuExit_Click()
  Call Unload(Me)
End Sub

Private Sub mnuMainMenu_Click()
    mnuConvertAll.Enabled = LV.ListCount > 0
    mnuClear.Enabled = LV.ListCount > 0
End Sub

Sub mnuOperations_Click(index As Integer)
  Dim i As Long
  Dim lLastItem As Long
  lLastItem = -1
  For i = LV.ListCount - 1 To 0 Step -1
    If LV.Selected(i) Then
      lLastItem = i
      Call LV.RemoveItem(i)
    End If
  Next
  If lLastItem >= 0 Then
    If lLastItem >= LV.ListCount Then lLastItem = LV.ListCount - 1
    If LV.ListCount > 0 Then
      LV.Selected(lLastItem) = True
    End If
  End If
  Call BackUpList
End Sub

Private Sub mnuOptions_Click()
  Call fOptions.Show(vbModal, Me)
End Sub

Private Sub picPanel_Resize()
    TabStrip1.Width = picPanel.ScaleWidth - TabStrip1.Left * 2
    Dim pic As PictureBox
    For Each pic In picFormats
        pic.Top = 400
        pic.Width = picPanel.ScaleWidth - pic.Left * 2
        pic.Height = 975 + 80
    Next
End Sub

Private Sub sldOGGQuality_Scroll(ByVal value As Double)
    Settings.OGG_Quality = Round(sldOGGQuality.value * 10, 2)
    lblQGGQualityLabel.Caption = Settings.OGG_Quality
End Sub

Private Sub TabStrip1_Click()
    Dim i As Long
    Dim index As Long
    index = TabStrip1.ListIndex
    Settings.ConvertToFormat = index
    For i = 0 To picFormats.Count - 1
        picFormats(i).Visible = CBool(i = index)
    Next
End Sub

Private Sub tmrCB_Timer()
  'Picture1.AutoRedraw = True
  'Dim i&
  'Picture1.Cls
  'For i = 0 To LV.ListCount - 1
  '  Picture1.Print LV.Focused(i)
  'Next
  On Error Resume Next
  Dim Tmp As String
  Dim lRet As Long
  Tmp = Clipboard.GetText
  If Not Tmp Like "*.mp3" Then Exit Sub
  If Tmp = OldCB Then Exit Sub
  Call AddFile(Tmp)
  OldCB = Tmp
  Call BackUpList
End Sub

Sub BackUpList()
  On Error Resume Next
  Dim i As Long
  Dim j As Long
  Dim FF As Long
  FF = FreeFile
  FileCopy RCP(App.Path) & "backup.bak", RCP(App.Path) & "backup_old.bak"
  Open RCP(App.Path) & "backup.bak" For Output As #FF
    For i = LV.ListCount - 1 To 0 Step -1
      For j = 0 To LV.ColumnCount - 1
        Print #FF, LV.SubItemText(i, j);
        If j < LV.ColumnCount - 1 Then
          Print #FF, vbTab;
        Else
          Print #FF, vbNullString
        End If
      Next
    Next
  Close #FF
  If LV.ListCount > 0 Then
    Caption = App.Title & " - " & CStr(LV.ListCount)
  Else
    Caption = App.Title
  End If
End Sub

Sub RestoreList()
  On Error Resume Next
  Dim i As Long
  Dim j As Long
  Dim FF As Long
  Dim Tmp As String
  Dim z  As Long
  Dim File As String
  FF = FreeFile
  Open RCP(App.Path) & "backup.bak" For Input As #FF
    Do While Not EOF(FF)
      Line Input #FF, Tmp
      If Tmp = vbNullString Then Exit Sub
      If InStr(Tmp, vbTab) > 0 Then
        File = Left(Tmp, InStr(Tmp, vbTab) - 1)
        Call AddFile(File)
      End If
    Loop
  Close #FF
End Sub

Private Sub LoadSettings()
    Dim dwACCPlusBitrate As Long
    sldOGGQuality.value = Settings.OGG_Quality / 10
    lblQGGQualityLabel = Round(sldOGGQuality.value * 10, 2)
    optBpp(0) = Settings.MP3_BitrateType = 0
    optBpp(1) = Settings.MP3_BitrateType = 1
    cmbVBpp.ListIndex = getListIndexByText(cmbVBpp, Settings.MP3_VariableBitrate)
    cmbACCPlusBr.ListIndex = getListIndexByText(cmbACCPlusBr, Settings.ACCPlus_Bitrate)
    cmbBpp.TEXT = Settings.MP3_Bitrate
    Call optBpp_Click(0)
End Sub

Function getListIndexByText(Obj As ComboBox, ByVal value)
    Dim i As Long
    For i = 0 To Obj.ListCount - 1
        If (Obj.List(i) = value) Then
            getListIndexByText = i
            Exit Function
        End If
    Next
End Function

Private Sub optBpp_Click(index As Integer)
    Settings.MP3_BitrateType = index
End Sub
