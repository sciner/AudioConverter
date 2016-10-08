Attribute VB_Name = "mMain"
Option Explicit

Public Enum ENCODE_TO_FORMAT_ENUM
    E2_WAV = 0
    E2_ACCPLUS = 1
    E2_OGG = 2
    E2_MP3 = 3
End Enum

Public Settings As New cSettings
Public CONFIG_FILE As String

Public Const DEFAULT_ACCP_BITRATE As Long = 32
Public Const KTC_FACE_COLOR As Long = vbButtonFace
Public Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Const WS_BORDER = &H800000
Private Const WS_POPUP = &H80000000
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260&
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER As Long = &H2000
Private Const GWL_EXSTYLE = (-20)
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Public Enum FindObjectType
  fNotFound = 0
  fFindFolder = 1
  fFindFile = 2
End Enum

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetInputState Lib "user32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal Section As String, ByVal Key As String, ByVal Default As String, ByVal GetStr As String, ByVal nSize As Long, ByVal INIfile As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Section As String, ByVal Key As String, ByVal putStr As String, ByVal INIfile As String) As Long

Public SupportFileTypes As New FileExtensions
Public GlobalCancel As Boolean
Public GlobalTimer As Long

'Необходимые нам API функции
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

'Константы
Const MIIM_TYPE = &H10
Const MIIM_SUBMENU = &H4
Const MIIM_BITMAP = &H80

Const MFT_RIGHTJUSTIFY = &H4000
Const MFT_STRING = &H0&
Const MF_MENUBARBREAK = &H20&
Const MF_BITMAP = &H4&

'Тип MENUITEMINFO
Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()
Dim CrcTableInit As Boolean
'Dim CRCTable(0 To 255) As Long

'Public Function CalcCRC32(bytearray() As Byte) As Long
'    Dim i As Long
'    Dim CRC As Long
'    If CrcTableInit = False Then Call Init_CRCTable
'    CRC = -1
'    For i = 0 To UBound(bytearray) - 1
'        CRC = (((CRC And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (CRCTable((CRC And &HFF) Xor bytearray(i)))
'    Next i
'    CRC = CRC Xor &HFFFFFFFF
'    CalcCRC32 = CRC
'End Function
'
'Private Sub Init_CRCTable()
'    Dim i As Long
'    Dim j As Long
'    Dim Limit As Long
'    Dim CRC As Long
'    Limit = &HEDB88320
'    For i = 0 To 255
'        CRC = i
'        For j = 0 To 7
'            If CRC And 1 Then
'              CRC = (((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor Limit
'            Else
'              CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
'            End If
'        Next j
'        CRCTable(i) = CRC
'    Next i
'    CrcTableInit = True
'End Sub

Function TrimNullChar(ByVal lpzString As String) As String
  lpzString = lpzString & vbNullChar
  lpzString = VBA.Left$(lpzString, InStr(lpzString, vbNullChar) - 1)
  TrimNullChar = VBA.Trim$(lpzString)
  End Function

Private Function GetSettingIni(ByVal File As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As Variant) As Variant
  Dim Tmp As String, ret As Long
  Tmp = VBA.Space$(16536)
  ret = GetPrivateProfileString(Section, Key, vbNullString, Tmp, Len(Tmp), File)
  If ret = 0 Then GetSettingIni = Default Else GetSettingIni = TrimNullChar(Tmp)
  End Function

Private Sub SaveSettingIni(ByVal File As String, ByVal Section As String, ByVal Key As String, ByVal value As Variant)
  Dim ret As Integer, ws As String
  Err.Clear
  ret = WritePrivateProfileString(Section, Key, CStr(value), File)
  End Sub

Public Property Get SETS(ByVal Key As String, Optional ByVal Default As Variant) As String
  On Error GoTo 1
  Dim Tmp As String
  Dim TempKey As String
  TempKey = Key
  Tmp = GetSettingIni(CONFIG_FILE, "OPTIONS", Key, Default)
  Tmp = Replace(Tmp, "Error 448", vbNullString)
  SETS = Tmp
  Exit Property
1:
  SETS = CStr(Default)
  End Property

Public Property Let SETS(ByVal Key As String, Optional ByVal Default As Variant, ByVal vNewValue As String)
  SaveSettingIni CONFIG_FILE, "OPTIONS", Key, vNewValue
  End Property

Sub Main()
    InitCommonControls
    Call fMain.Show
End Sub

Function FormIsLoad(ByVal FName$) As Boolean
  Dim i&
  For i = 0 To Forms.Count - 1
    If Forms(i).Name = FName Then
      FormIsLoad = True
      Exit Function
    End If
  Next
End Function

Sub ShiftMenu(ByVal hWnd As Long, ByVal index As Long, ByVal sCaption As String)
  Dim MnuInfo As MENUITEMINFO
  Dim mnuH As Long
  Dim MyTemp As Long
  mnuH = GetMenu(hWnd)
  MyTemp = GetMenuItemInfo(mnuH, index, True, MnuInfo)
  With MnuInfo
    .cbSize = Len(MnuInfo)
    .fMask = MIIM_TYPE
    .fType = MFT_RIGHTJUSTIFY Or MFT_STRING
    .cch = Len(sCaption)
    .dwTypeData = sCaption
    .cbSize = Len(MnuInfo)
  End With
  MyTemp = SetMenuItemInfo(mnuH, index, True, MnuInfo)
  MyTemp = DrawMenuBar(hWnd)
End Sub

'Если путь [Не найден] возвращает 0
'Если путь [Папка] возвращает 1
'Если путь [Файл] возвращает 2
Function Find(ByVal Path As String) As FindObjectType
  Dim lRet As Long
  Dim W32 As WIN32_FIND_DATA
  If VBA.Right$(Path, 1) = "\" Then Path = VBA.Left$(Path, Len(Path) - 1)
  lRet = FindFirstFile(Path, W32)
  If lRet = INVALID_HANDLE_VALUE Then Exit Function
  Call FindClose(lRet)
  Find = IIf(W32.dwFileAttributes And vbDirectory, fFindFolder, fFindFile)
End Function

Sub OnlyNumbers(TBox As TextBox)
  Call SetWindowLong(TBox.hWnd, GWL_STYLE, GetWindowLong&(TBox.hWnd, GWL_STYLE) Or ES_NUMBER)
  Call TBox.Refresh
End Sub

Sub ChangeWindowStyle(ByVal hWnd As Long)
  On Local Error Resume Next
  Dim TFlat As Long
  TFlat = GetWindowLong(hWnd, GWL_EXSTYLE)
  TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  SetWindowLong hWnd, GWL_EXSTYLE, TFlat
  SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

'added by SCINER 13/01/2006 4:36 ***********************
  'Определяет запущена ли программа из под IDE
  Function InIDE() As Boolean
    Debug.Assert SetTrue(InIDE)
  End Function
  Private Function SetTrue(bValue As Boolean) As Boolean
    SetTrue = True
    bValue = True
  End Function
'******************************************************

Property Get CopyToPath() As String
    Dim lRet As Long
    CopyToPath = Settings.Save_Directory
2:
    If Find(CopyToPath) <> fFindFolder Then
      lRet = fSelectFolder.GetOption
      Select Case lRet
      Case 0 'manual select
        CopyToPath = BrowseFolders(fMain.hWnd, "Выберите папку, куда копировать файлы:", BrowseForFolders, CSIDL_DESKTOP) '+весь компьютер
        GoTo 2
      Case 1 'default
        CopyToPath = App.Path
      Case 2 'exit
        Unload fSelectFolder
        Set fSelectFolder = Nothing
        Unload fMain
      End Select
      Settings.Save_Directory = RCP(CopyToPath)
    End If
End Property

'Public Function Replace(ByVal sExpression As String, sFind As String, sReplace As String, Optional vStart As Long = 1, Optional vCount As Long = -1, Optional vCompare As Long = vbBinaryCompare) As String
'  Dim i As Long, L As Long, k As Long, H As Long
'  L = Len(sFind)
'  If vStart > Len(sExpression) Then Exit Function
'  If L > 0 And L <= Len(sExpression) Then
'  k = Len(sReplace) '+ 1
'  If vCount < 0 Then vCount = 2147483647
'  For H = 1 To vCount
'    i = InStr(vStart, sExpression, sFind, vCompare)
'    If i = 0 Then Exit For
'    sExpression = VBA.Left$(sExpression, i - 1) & sReplace _
'    & VBA.Mid$(sExpression, i + L): vStart = i + k
'  Next H
'  End If
'  Replace = sExpression
'End Function
'
'Function SplitVB5(Tp, ByVal Tmp$, ByVal Splt$) As Long
'  Dim i&
'  Dim j As Long
'  i = InStr(Tmp, Splt)
'  Do While i > 0
'    ReDim Preserve Tp(j)
'    Tp(j) = VBA.Left$(Tmp, i - 1)
'    Tmp = VBA.Mid$(Tmp, i + Len(Splt))
'    j = j + 1
'    i = InStr(Tmp, Splt)
'    DoEvents
'  Loop
'  ReDim Preserve Tp(j)
'  Tp(j) = Tmp
'  j = j + 1
'  SplitVB5 = j
'End Function
'
Function Trunc(ByVal X, ByVal Min, ByVal Max)
  If X < Min Then X = Min
  If X > Max Then X = Max
  Trunc = X
End Function
