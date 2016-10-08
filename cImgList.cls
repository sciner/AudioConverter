VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImgList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Created by SCINER: lenar2003@mail.ru

Private Const MAX_PATH = 260

Private Const ILC_MASK = &H1
Private Const ILC_COLOR = &H0
Private Const ILC_COLORDDB = &HFE
Private Const ILC_COLOR4 = &H4
Private Const ILC_COLOR8 = &H8
Private Const ILC_COLOR16 = &H10
Private Const ILC_COLOR24 = &H18
Private Const ILC_COLOR32 = &H20

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32" ()

Private Const SHGFI_ICON = &H100                         '  get icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index

Private Type SHFILEINFO
  hIcon As Long                      '  out: icon
  iIcon As Long          '  out: icon index
  dwAttributes As Long               '  out: SFGAO_ flags
  szDisplayName As String * MAX_PATH '  out: display name (or path)
  szTypeName As String * 80         '  out: type name
End Type

Private Declare Function ImageList_Create Lib "COMCTL32.DLL" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "COMCTL32.DLL" (ByVal himl As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "COMCTL32.DLL" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_AddIcon Lib "COMCTL32.DLL" (ByVal hImageList As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" (ByVal himl As Long) As Long
Private Declare Function ImageList_SetImageCount Lib "COMCTL32.DLL" (ByVal himl As Long, ByVal uNewCount As Long) As Long

Private hwIL As Long
Dim NullIcn As Long

Property Get NullIcon(ByVal h As Long)
  NullIcn = h
End Property

Property Get hWnd() As Long
  hWnd = hwIL
End Property

Private Sub CreateMe()
  hwIL = ImageList_Create(16, 16, ILC_COLOR8 Or ILC_MASK, 1, 1)
End Sub

Sub Clear()
  ImageList_SetImageCount hwIL, 0
  'Call Class_Terminate
  'hwIL = 0
End Sub

Private Sub Class_Initialize()
  InitCommonControls
  Call CreateMe
End Sub

Private Sub Class_Terminate()
  Call ImageList_Destroy(hwIL)
End Sub

Sub AddIcon(ByVal hIcon As Long)
  'If hwIL = 0 Then CreateMe
  Call ImageList_ReplaceIcon(hwIL, -1, hIcon)
End Sub

Function AddIconFromFile(ByVal Path As String) As Long

  'If hwIL = 0 Then CreateMe
  
  Dim hIcon As Long
  Dim shInfo As SHFILEINFO
  SHGetFileInfo Path, 0, shInfo, Len(shInfo), SHGFI_ICON Or SHGFI_SMALLICON
  If shInfo.hIcon = 0 Then shInfo.hIcon = NullIcn
  hIcon = shInfo.hIcon

  AddIcon hIcon
  DestroyIcon hIcon
  AddIconFromFile = ImageList_GetImageCount(hwIL) - 1

End Function
