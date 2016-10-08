VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MP3_BitrateType As Long
Public MP3_VariableBitrate As Long
Public MP3_Bitrate As Integer

Public ACCPlus_Bitrate As Long

Public OGG_Quality As Double

Public Clipboard_Monitoring As Boolean
Public SaveTo_Method As Long
Public Save_Directory As String
Public Replace_Original_File As Boolean

Public ConvertToFormat As Long

Public Sub LoadSettings()
    MP3_Bitrate = SETS("MP3Bpp", 128)
    MP3_BitrateType = SETS("MP3BitrateType", 0)
    MP3_VariableBitrate = SETS("MP3vBpp", 128)
    ACCPlus_Bitrate = val(SETS("ACCPlusBpp", DEFAULT_ACCP_BITRATE))
    OGG_Quality = val(Replace(SETS("OGGQuality", 50), ",", "."))
    Replace_Original_File = SETS("ReplaceOriginal", 0) = "1"
    SaveTo_Method = val(SETS("SaveTo_Method", 1))
    Clipboard_Monitoring = SETS("Clipboard_Monitoring", 0) = "1"
    Save_Directory = RCP(SETS("Save_Directory", RCP(App.Path)))
    ConvertToFormat = val(SETS("ConvertToFormat", 2))
    Call Normalize
End Sub

Private Sub Normalize()
    OGG_Quality = Trunc(OGG_Quality, 0, 10)
    ACCPlus_Bitrate = Int(Trunc(ACCPlus_Bitrate, 16, 256))
    SaveTo_Method = Int(Trunc(SaveTo_Method, 0, 1))
    ConvertToFormat = Int(Trunc(ConvertToFormat, 0, 3))
    Save_Directory = RCP(Save_Directory)
End Sub

Public Sub SaveSettings()
    Call Normalize
    SETS("MP3Bpp") = MP3_Bitrate
    SETS("MP3BitrateType") = MP3_BitrateType
    SETS("MP3vBpp") = MP3_VariableBitrate
    SETS("ACCPlusBpp") = ACCPlus_Bitrate
    SETS("OGGQuality") = OGG_Quality
    SETS("ReplaceOriginal") = IIf(Replace_Original_File, 1, 0)
    SETS("Clipboard_Monitoring") = IIf(Clipboard_Monitoring, 1, 0)
    SETS("SaveTo_Method") = SaveTo_Method
    SETS("Save_Directory") = RCP(Save_Directory)
    SETS("ConvertToFormat") = ConvertToFormat
End Sub

Public Property Get GetBpp() As Long
  GetBpp = val(SETS("MP3Bpp", "128"))
  If GetBpp Mod 2 <> 0 Then GetBpp = 128
  If InStr(",32,40,48,56,64,80,96,112,128,160,192,224,256,320,", "," & CStr(GetBpp) & ",") = 0 Then GetBpp = 128
End Property

Private Sub Class_Initialize()
    Call LoadSettings
End Sub
