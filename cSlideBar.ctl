VERSION 5.00
Begin VB.UserControl cSlideBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "cSlideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Created by SCINER: lenar2003@mail.ru
'SlideBar like PhotoShop
'14/01/2006 4:29
'For project LivePaint: by Sources.Ru Team

Event Scroll(ByVal value As Double)
Event Change(ByVal value As Double)

Dim val As Double
Dim W&, H&
Dim cW&
Const TrW = 10
Dim L&
Dim T&
Dim OldValue As Double

Sub Redraw()
  Dim i&
  Cls
  DrawWidth = 1
  T = 3
  L = TrW \ 2
  Line (L, T)-Step(cW, 0)
  PSet (L + cW / 1 * val, T + 1)
  For i = 0 To 10
    Line (L + cW / 1 * val - i \ 2, T + 1 + i)-Step(i, 0)
  Next
  'Line -Step(-TrW / 2, TrW)
  'Line -Step(TrW, 0)
  'Line -Step(-TrW / 2, -TrW)
  For i = 0 To 100 Step 10
    DrawWidth = IIf(i Mod 50 = 0, 2, 1)
    Line (L + (cW / 100 * i), T)-Step(0, -IIf((i Mod 10 = 0) Or (i Mod 50 = 0), T, T / 2))
  Next
End Sub

Public Property Get value() As Double
  value = val
End Property

Public Property Let value(ByVal vNewValue As Double)
  If val = vNewValue Then Exit Property
  val = vNewValue
  Call Redraw
  RaiseEvent Change(val)
End Property

Private Sub UserControl_Initialize()
  AutoRedraw = True
  ScaleMode = vbPixels
  Call Redraw
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> vbLeftButton Then Exit Sub
  If X > L + cW Then X = L + cW
  If X < L Then X = L
  X = X - L
  val = (X / cW * 1)
  RaiseEvent Scroll(val)
  Call Redraw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    If val <> OldValue Then
      RaiseEvent Change(val)
      OldValue = val
    End If
  End If
End Sub

Private Sub UserControl_Resize()
  W = ScaleWidth - 1
  H = ScaleHeight
  cW = W - TrW
  Call Redraw
End Sub
