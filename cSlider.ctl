VERSION 5.00
Begin VB.UserControl cSlider 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "cSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Created by SCINER: lenar2003@mail.ru

Event Change(ByVal Value As Double)
Private Const COLOR_BTNFACE = 15

'DrawEdge Constants
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT   As Long = &H1
Private Const BF_RIGHT  As Long = &H4
Private Const BF_TOP    As Long = &H2
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT   As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, pRect As RECT, ByVal lEdge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Dim DC&, Bmp&, Obj&, W&, H&
Dim lSldSizeW As Long
Dim lSldSizeH As Long
Dim lRailSize As Long
Dim RC As RECT
Dim lGrayBrush As Long
Dim lVal As Double
Dim lPreviousXPos As Long
Dim lMin As Double
Dim lMax As Double

Private Sub Redraw()
  With RC
    .Left = 0
    .Top = 0
    .Right = W
    .Bottom = H
  End With
  Call FillRect(DC, RC, lGrayBrush)
  With RC
    .Left = lSldSizeW \ 2
    .Top = H \ 2 - lRailSize \ 2
    .Right = W - lSldSizeW \ 2
    .Bottom = .Top + lRailSize
  End With
  Call DrawEdge(DC, RC, BDR_SUNKENOUTER, BF_RECT)
  With RC
    .Left = (W - lSldSizeW) / 100 * lVal
    .Top = H \ 2 - lSldSizeH \ 2
    .Right = .Left + lSldSizeW
    .Bottom = .Top + lSldSizeH
  End With
  Call FillRect(DC, RC, lGrayBrush)
  Call DrawEdge(DC, RC, EDGE_RAISED, BF_RECT)

  With RC
    .Left = (W - lSldSizeW) / 100 * lVal + 2
    .Top = H \ 2 - lSldSizeH \ 2 + 2
    .Right = .Left + lSldSizeW - 4
    .Bottom = .Top + lSldSizeH - 4
  End With
  Call DrawEdge(DC, RC, BDR_SUNKENOUTER, BF_RECT)
  Call FlipFlop
End Sub

Sub FlipFlop()
  Call BitBlt(hdc, 0, 0, W, H, DC, 0, 0, vbSrcCopy)
End Sub

'Создание заднего буфера
Sub CreateBackDC(ByVal ParentDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, _
                 lDc As Long, lBmp As Long, lObj As Long)
  lDc = CreateCompatibleDC(0)
  lBmp = CreateCompatibleBitmap(ParentDC, lWidth, lHeight)
  lObj = SelectObject(lDc, lBmp)
End Sub

Private Sub UserControl_Initialize()
  ScaleMode = vbPixels
  lSldSizeW = 14
  lSldSizeH = 14
  lRailSize = 5
  'lMin = 80
  'lMax = 50
  lGrayBrush = GetSysColorBrush(COLOR_BTNFACE)
End Sub

Private Sub CheckSizes()
  If W < lSldSizeW Then Width = lSldSizeW * Screen.TwipsPerPixelX
  If H < lSldSizeH Then Height = lSldSizeH * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> vbLeftButton Then Exit Sub
  If X < lSldSizeW \ 2 Then X = lSldSizeW \ 2
  If X > W - lSldSizeW \ 2 Then X = W - lSldSizeW \ 2
  X = X - lSldSizeW \ 2
  If X = lPreviousXPos Then Exit Sub
  lPreviousXPos = X
  lVal = X / (W - lSldSizeW) * 100
  Call Redraw
  RaiseEvent Change(lMin + (lMax - lMin) / 100 * lVal)
End Sub

Private Sub UserControl_Paint()
  Call FlipFlop
End Sub

Sub Destroy()
  If DC <> 0 Then DeleteDC DC
  If Bmp <> 0 Then DeleteObject Bmp
  If Obj <> 0 Then DeleteObject Obj
End Sub


Private Sub UserControl_Resize()
  W = ScaleWidth
  H = ScaleHeight
  Call Destroy
  Call CreateBackDC(hdc, W, H, DC, Bmp, Obj)
  Call CheckSizes
  Call Redraw
End Sub

Private Sub UserControl_Terminate()
  Call Destroy
End Sub

Public Property Get Min() As Double
  Min = lMin
End Property
Public Property Let Min(ByVal vNewValue As Double)
  lMin = vNewValue
End Property
Public Property Get Max() As Double
  Max = lMax
End Property
Public Property Let Max(ByVal vNewValue As Double)
  lMax = vNewValue
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  lMin = PropBag.ReadProperty("Min", 0)
  lMax = PropBag.ReadProperty("Max", 100)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Min", lMin, 0)
  Call PropBag.WriteProperty("Max", lMax, 100)
End Sub

Public Property Get Value() As Double
  Value = lMin + (lMax - lMin) / 100 * lVal
End Property

Public Property Let Value(ByVal vNewValue As Double)
  lVal = (((vNewValue - Min)) / (Max - Min)) * 100
  Call Redraw
  RaiseEvent Change(Value)
End Property


