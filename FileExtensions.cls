VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Created by SCINER: lenar2003@mail.ru
'Класс управления расширениями файлов

Dim exf As Collection
Public Function Exists(ByVal Path As String) As Boolean
  Dim i As Long
  Path = VBA.LCase$(Path)
  For i = 0 To Count - 1
    If Path Like "*." & Item(i) Then
      Exists = True
      Exit Function
    End If
  Next
End Function
Sub Clear()
  Set exf = New Collection
End Sub
Property Get Count() As Long
  Count = exf.Count
End Property
Property Get Item(ByVal index As Long) As String
  Item = exf.Item(index + 1)
End Property
Sub Add(ByVal ext As String)
  Call exf.Add(VBA.LCase$(ext))
End Sub
Private Sub Class_Initialize()
  Call Clear
End Sub
Function GetForShowDialog() As String
  Dim Tmp As String
  Dim i As Long
  For i = 0 To Count - 1
    Tmp = Tmp & "*." & Item(i) & "|*." & Item(i) & "|"
  Next
  If Len(Tmp) > 0 Then Tmp = VBA.Left$(Tmp, Len(Tmp) - 1)
  GetForShowDialog = Tmp
End Function
Sub FillListBox(LBox As ListBox)
  Call LBox.Clear
  Dim i As Long
  For i = 0 To Count - 1
    Call LBox.AddItem("*." & Item(i))
  Next
End Sub
