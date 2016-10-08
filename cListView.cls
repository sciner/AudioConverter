VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Created by SCINER: lenar2003@mail.ru

'WM_CONTEXTMENU
Dim m_HideSelection As Boolean

Private Const LVS_SHOWSELALWAYS        As Long = &H8
Private Const WM_SETFOCUS = &H7

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Enum ItemStateEnum
  LVIS_FOCUSED = &H1
  LVIS_SELECTED = &H2
  LVIS_CUT = &H4
  LVIS_DROPHILITED = &H8
  LVIS_ACTIVATING = &H20
End Enum

Private Const LVM_FIRST                    As Long = &H1000
Private Const LVM_GETBKCOLOR               As Long = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR               As Long = (LVM_FIRST + 1)
Private Const LVM_GETIMAGELIST             As Long = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST             As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT             As Long = (LVM_FIRST + 4)
Private Const LVM_GETITEM                  As Long = (LVM_FIRST + 5)
Private Const LVM_SETITEM                  As Long = (LVM_FIRST + 6)
Private Const LVM_INSERTITEM               As Long = (LVM_FIRST + 7)
Private Const LVM_DELETEITEM               As Long = (LVM_FIRST + 8)
Private Const LVM_DELETEALLITEMS           As Long = (LVM_FIRST + 9)
Private Const LVM_GETNEXTITEM              As Long = (LVM_FIRST + 12)
Private Const LVM_FINDITEM                 As Long = (LVM_FIRST + 13)
Private Const LVM_HITTEST                  As Long = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE            As Long = (LVM_FIRST + 19)
Private Const LVM_SCROLL                   As Long = (LVM_FIRST + 20)
Private Const LVM_REDRAWITEMS              As Long = (LVM_FIRST + 21)
Private Const LVM_ARRANGE                  As Long = (LVM_FIRST + 22)
Private Const LVM_EDITLABEL                As Long = (LVM_FIRST + 23)
Private Const LVM_GETEDITCONTROL           As Long = (LVM_FIRST + 24)
Private Const LVM_GETCOLUMN                As Long = (LVM_FIRST + 25)
Private Const LVM_SETCOLUMN                As Long = (LVM_FIRST + 26)
Private Const LVM_INSERTCOLUMN             As Long = (LVM_FIRST + 27)
Private Const LVM_DELETECOLUMN             As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH           As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH           As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER                As Long = (LVM_FIRST + 31)
Private Const LVM_GETTEXTCOLOR             As Long = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR             As Long = (LVM_FIRST + 36)
Private Const LVM_GETTEXTBKCOLOR           As Long = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR           As Long = (LVM_FIRST + 38)
Private Const LVM_UPDATE                   As Long = (LVM_FIRST + 42)
Private Const LVM_SETITEMSTATE             As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE             As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT              As Long = (LVM_FIRST + 45)
Private Const LVM_SETITEMTEXT              As Long = (LVM_FIRST + 46)
Private Const LVM_SORTITEMS                As Long = (LVM_FIRST + 48)
Private Const LVM_GETSELECTEDCOUNT         As Long = (LVM_FIRST + 50)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Private Const LVM_SETHOTITEM               As Long = (LVM_FIRST + 60)
Private Const LVM_GETHOTITEM               As Long = (LVM_FIRST + 61)
Private Const LVM_SETHOTCURSOR             As Long = (LVM_FIRST + 62)
Private Const LVM_GETHOTCURSOR             As Long = (LVM_FIRST + 63)
Private Const LVM_SETBKIMAGE               As Long = (LVM_FIRST + 68)
Private Const LVM_GETBKIMAGE               As Long = (LVM_FIRST + 69)
Private Const LVM_SORTITEMSEX              As Long = (LVM_FIRST + 81)
Private Const LVM_SETVIEW                  As Long = (LVM_FIRST + 142)
Private Const LVM_GETVIEW                  As Long = (LVM_FIRST + 143)

Private Const LVIF_TEXT           As Long = &H1
Private Const LVIF_IMAGE          As Long = &H2
Private Const LVIF_PARAM          As Long = &H4
Private Const LVIF_STATE          As Long = &H8
Private Const LVIF_INDENT         As Long = &H10
Private Const LVIF_GROUPID        As Long = &H100
Private Const LVIF_COLUMNS        As Long = &H200

Public Enum eSortAsConst
  stText = 0
  stIndex = 1
  stDate = 2
  stNumber = 3
End Enum

Public Enum eSortOrderConstants
    '[soDefault] = 0
    [soAscending] = 1
    [soDescending] = -1
End Enum

Public Enum AlignConsts
  LVCFMT_LEFT = &H0
  LVCFMT_RIGHT = &H1
  LVCFMT_CENTER = &H2
End Enum

Private Type LVITEM_lp
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    stateMask  As Long
    pszText    As Long
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type

Private Enum LVENUMS
  LVCF_FMT = &H1
  LVCF_WIDTH = &H2
  LVCF_TEXT = &H4
  LVCF_SUBITEM = &H8
  LVCF_IMAGE = &H10
  LVCF_ORDER = &H20
  LVCFMT_JUSTIFYMASK = &H3
  LVCFMT_IMAGE = &H800
  LVCFMT_BITMAP_ON_RIGHT = &H1000
  LVCFMT_COL_HAS_IMAGES = &H8000
End Enum

Private Type LV_COLUMN
  mask As Long
  fmt As Long
  cx As Long
  pszText As Long
  cchTextMax As Long
  iSubItem As Long
  iImage As Long
  iOrder As Long
End Type

Public Enum eListViewStyle
  LVS_ICON = &H0
  LVS_REPORT = &H1
  LVS_SMALLICON = &H2
  LVS_LIST = &H3
End Enum

Private Const GWL_STYLE = (-16)

' LVM_GET/SETIMAGELIST wParam
Public Enum ImageListSize
  LVSIL_NORMAL = 0
  LVSIL_SMALL = 1
End Enum

Private Const LVM_SETEXTENDEDSTYLE      As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDSTYLE      As Long = (LVM_FIRST + 55)

Private Const WM_USER = &H400
Private Const IDC_LISTVIEW = WM_USER + 2130
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_CHILD = &H40000000

Private Const WC_LISTVIEW = "SysListView32"

Private Const LVS_EX_GRIDLINES          As Long = &H1&
Private Const LVS_EX_DOUBLEBUFFER       As Long = &H10000
Private Const LVS_EX_FLATSB             As Long = &H100&
Private Const LVS_EX_SUBITEMIMAGES      As Long = &H2&
Private Const LVS_EX_CHECKBOXES         As Long = &H4&
Private Const LVS_EX_INFOTIP            As Long = &H400&
Private Const LVS_EX_LABELTIP           As Long = &H4000&
Private Const LVS_EX_TRACKSELECT        As Long = &H8&
Private Const LVS_EX_HEADERDRAGDROP     As Long = &H10&
Private Const LVS_EX_FULLROWSELECT      As Long = &H20&
Private Const LVS_EX_ONECLICKACTIVATE   As Long = &H40&
Private Const LVS_EX_BORDERSELECT       As Long = &H8000&
Private Const LVS_EX_UNDERLINEHOT       As Long = &H800&

Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2 'Note: On last column, its width fills remaining width

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32" ()

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_FRAMECHANGED  As Long = &H20

Private Type MyLVType
  hWnd As Long
  View As eListViewStyle
  ColumnCount  As Long
  style As Long
End Type

Dim LView As MyLVType

Private Type LVITEM
  mask       As Long
  iItem      As Long
  iSubItem   As Long
  State      As Long
  stateMask  As Long
  pszText    As String
  cchTextMax As Long
  iImage     As Long
  lParam     As Long
  iIndent    As Long
End Type

Private Sub Class_Initialize()
  Call InitCommonControls
End Sub

Sub Sort(Optional ByVal ColumnIndex As Long = 0, _
         Optional ByVal SortAs As eSortAsConst = stText, _
         Optional ByVal SortOrder As eSortOrderConstants = [soAscending])
  m_lColumn = ColumnIndex
  lSortAs = SortAs
  m_PRECEDE = SortOrder
  m_FOLLOW = -SortOrder
  SendMessageLong LView.hWnd, LVM_SORTITEMSEX, LView.hWnd, AddressOf hSortFunc
End Sub

Property Get ListCount() As Long
  ListCount = SendMessage(LView.hWnd, LVM_GETITEMCOUNT, 0, ByVal 0&)
End Property

Public Function ItemAdd(ByVal Item As Integer, ByVal TEXT As String, ByVal indent As Integer, ByVal Icon As Integer) As Long
  Dim uLV As LVITEM
  With uLV
      .iItem = Item
      .lParam = Item
      .pszText = TEXT & vbNullChar
      .cchTextMax = Len(TEXT) + 1
      .iIndent = indent
      .iImage = Icon
      .mask = LVIF_TEXT Or LVIF_INDENT Or LVIF_IMAGE Or LVIF_PARAM
  End With
  ItemAdd = SendMessage(LView.hWnd, LVM_INSERTITEM, 0, uLV) ' (SendMessage(LView.hWnd, LVM_INSERTITEM, 0, uLV) > -1)
End Function

Public Property Get SubItemText(ByVal Item As Integer, ByVal SubItem As Integer) As String
  Dim uLVI As LVITEM_lp
  Dim a(256) As Byte
  Dim lLen As Long
    If (LView.hWnd) Then
      With uLVI
        .iSubItem = SubItem
        .pszText = VarPtr(a(0))
        .cchTextMax = UBound(a)
        .mask = LVIF_TEXT
      End With
      lLen = SendMessage(LView.hWnd, LVM_GETITEMTEXT, Item, uLVI)
      SubItemText = VBA.Left$(StrConv(a(), vbUnicode), lLen)
    End If
End Property
Public Property Let SubItemText(ByVal Item As Integer, ByVal SubItem As Integer, ByVal TEXT As String)
  Dim uLVI As LVITEM
  If (LView.hWnd) Then
    With uLVI
      .iSubItem = SubItem
      .pszText = TEXT & vbNullChar
      .cchTextMax = Len(TEXT) + 1
    End With
    Call SendMessage(LView.hWnd, LVM_SETITEMTEXT, Item, uLVI)
  End If
End Property


Sub AutoSizeColumn(ByVal index As Long)
  SendMessage LView.hWnd, LVM_SETCOLUMNWIDTH, index, LVSCW_AUTOSIZE_USEHEADER
End Sub

Private Sub pSetExtendedStyle(ByVal iStyle As Long, ByVal bNew As Boolean)
  If bNew Then LView.style = LView.style Or iStyle Else LView.style = LView.style And Not iStyle
  If bNew Then
    SendMessage LView.hWnd, LVM_SETEXTENDEDSTYLE, 0, ByVal SendMessage(LView.hWnd, LVM_GETEXTENDEDSTYLE, 0, 0) Or iStyle
  Else
    SendMessage LView.hWnd, LVM_SETEXTENDEDSTYLE, 0, ByVal SendMessage(LView.hWnd, LVM_GETEXTENDEDSTYLE, 0, 0) And Not iStyle
  End If
End Sub

Property Let FullRowSelect(bol As Boolean)
  pSetExtendedStyle LVS_EX_FULLROWSELECT, bol
End Property
Property Get FullRowSelect() As Boolean
  FullRowSelect = LView.style And LVS_EX_FULLROWSELECT
End Property

Property Let FlatScrollBar(bol As Boolean)
  pSetExtendedStyle LVS_EX_FLATSB, bol
End Property
Property Get FlatScrollBar() As Boolean
  FlatScrollBar = LView.style And LVS_EX_FLATSB
End Property

Property Let GridLines(bol As Boolean)
  pSetExtendedStyle LVS_EX_GRIDLINES, bol
End Property
Property Get GridLines() As Boolean
  GridLines = LView.style And LVS_EX_GRIDLINES
End Property

Property Get ColumnCount() As Long
  ColumnCount = LView.ColumnCount
End Property

Function ColumnAdd(ByVal Caption As String, ByVal Width As Long, Optional ByVal Align As AlignConsts = LVCFMT_LEFT) As Boolean
  Dim LVC As LV_COLUMN
  LVC.mask = LVCF_FMT Or LVCF_TEXT Or LVCF_WIDTH Or LVCF_SUBITEM
  LVC.pszText = StrPtr(StrConv(Caption, vbFromUnicode))
  LVC.fmt = Align
  LVC.cx = Width
  LVC.iOrder = 0
  Call SendMessage(LView.hWnd, LVM_INSERTCOLUMN, LView.ColumnCount, LVC)
  LView.ColumnCount = LView.ColumnCount + 1
End Function

Function Init(ByVal hWndParent As Long) As Boolean
  LView.hWnd = CreateWindowEx(WS_EX_CLIENTEDGE, WC_LISTVIEW, vbNullString, WS_CHILD, 0, 0, 200, 100, hWndParent, 0&, App.hInstance, 0&)
  If LView.hWnd = 0 Then Exit Function
  Init = True
  Call ShowWindow(LView.hWnd, 1)
  Call HookListview(LView.hWnd)
End Function

Property Set ImageList(ImageSize As ImageListSize, ImgList As cImgList)
  Call SendMessage(LView.hWnd, LVM_SETIMAGELIST, ImageSize, ByVal ImgList.hWnd)
End Property

Sub Move(Left As Single, Optional Top As Single, Optional Width As Single, Optional Height As Single)
  MoveWindow LView.hWnd, Left, Top, Width, Height, 1
End Sub

Property Let Visible(vis As Boolean)
  ShowWindow LView.hWnd, IIf(vis, 1, 0)
End Property

Public Property Get View() As eListViewStyle
  View = LView.View
End Property

Sub Clear()
  SendMessage LView.hWnd, LVM_DELETEALLITEMS, 0, ByVal 0&
End Sub

Public Property Let View(ByVal vNewValue As eListViewStyle)
  Dim iNot As Long
  iNot = (LVS_ICON Or LVS_SMALLICON Or LVS_REPORT Or LVS_LIST)
  SetWindowLong LView.hWnd, GWL_STYLE, (GetWindowLong(LView.hWnd, GWL_STYLE) And Not iNot) Or vNewValue
End Property

Public Property Get hWnd() As Long
  hWnd = LView.hWnd
End Property

Public Property Get Focused(ByVal index As Long) As Boolean
  Dim iState As Long
  iState = LVIS_FOCUSED
  Focused = CBool(SendMessageLong(LView.hWnd, LVM_GETITEMSTATE, index, iState))
End Property

Private Property Let ItemSelected(ByVal Item As Integer, ByVal Selected As Boolean)
  Dim uLVI As LVITEM
  With uLVI
      .stateMask = LVIS_SELECTED Or -(Selected And Item > -1) * LVIS_FOCUSED
      .State = -Selected * LVIS_SELECTED Or -(Item > -1) * LVIS_FOCUSED
      .mask = LVIF_STATE
  End With
  Call SendMessage(LView.hWnd, LVM_SETITEMSTATE, Item, uLVI)
End Property

Public Property Let Selected(ByVal index As Long, ByVal bState As Boolean)
  ItemSelected(index) = bState
  SendMessage LView.hWnd, WM_SETFOCUS, 0, 0
End Property

Public Property Get Selected(ByVal index As Long) As Boolean
  Dim iState As Long
  iState = LVIS_SELECTED 'Or LVIS_FOCUSED
  Selected = CBool(SendMessageLong(LView.hWnd, LVM_GETITEMSTATE, index, iState))
End Property

Sub RemoveItem(ByVal index As Long)
  Call SendMessage(LView.hWnd, LVM_DELETEITEM, index, 0)
  'Returns TRUE if successful, or FALSE otherwise.
End Sub

Public Property Get HideSelection() As Boolean
    HideSelection = m_HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
  m_HideSelection = New_HideSelection
  If (m_HideSelection) Then
      Call pvSetStyle(0, LVS_SHOWSELALWAYS)
    Else
      Call pvSetStyle(LVS_SHOWSELALWAYS, 0)
  End If
End Property

Private Sub pvSetStyle(ByVal lStyle As Long, ByVal lStyleNot As Long)
  Dim lS As Long
    If (LView.hWnd) Then
        lS = GetWindowLong(LView.hWnd, GWL_STYLE)
        lS = lS And Not lStyleNot
        lS = lS Or lStyle
        Call SetWindowLong(LView.hWnd, GWL_STYLE, lS)
        Call SetWindowPos(LView.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
   End If
End Sub
