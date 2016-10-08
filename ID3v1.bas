Attribute VB_Name = "ID3v1"
Option Explicit

Public Function GetID3v1(Optional ByVal Path As String = vbNullString) As String
  On Error GoTo 2
  If Path = vbNullString Then GoTo 2
  Dim FF As Long
  Dim TAG$
  FF = FreeFile
  Open Path For Binary As #FF
    TAG = VBA.Space$(128)
    Get #FF, LOF(FF) - 127, TAG
  Close #FF
2:
  If VBA.Left$(TAG, 3) <> "TAG" Then
    TAG = VBA.String$(128, vbNullChar)
    Mid(TAG, 1, 3) = "TAG"
    Mid(TAG, 4, 30) = VBA.Left$(MusicName, 30)
    Mid(TAG, 94, 4) = Year(Now)
    Mid(TAG, 98, 17) = "lenar2003@mail.ru"
  End If
  GetID3v1 = TAG
End Function

Public Sub SetID3v1(ByVal Path As String, ByVal TAG As String)
  Dim FF As Long
  If VBA.Left(TAG, 3) <> "TAG" Then Exit Sub
  FF = FreeFile
  Open Path For Binary As #FF
    If LOF(FF) > 0 Then Put #FF, LOF(FF), TAG
  Close #FF
End Sub
