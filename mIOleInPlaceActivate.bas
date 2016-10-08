Attribute VB_Name = "mIOleInPlaceActivate"
'========================================================================================
' Filename:    mIOleInPlaceActivate.bas
' Author:      Mike Gainer, Matt Curland and Bill Storage
' Date:        09 January 1999
'
' Requires:    OleGuids.tlb (in IDE only)
'
' Description:
' Allows you to replace the standard IOLEInPlaceActiveObject interface for a
' UserControl with a customisable one.  This allows you to take control
' of focus in VB controls.
'
' The code could be adapted to replace other UserControl OLE interfaces.
'
' ---------------------------------------------------------------------------------------
' Visit vbAccelerator, advanced, free source for VB programmers
' http://vbaccelerator.com
'========================================================================================

Option Explicit

Private Const WM_KEYUP = &H101
Private Const WM_KEYDOWN = &H100
Public Const WM_MOUSEACTIVATE = &H21

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'========================================================================================
' Lightweight object definition
'========================================================================================

Public Type IPAOHookStructListView
    lpVTable    As Long                    'VTable pointer
    IPAOReal    As IOleInPlaceActiveObject 'Un-AddRefed pointer for forwarding calls
    Ctl         As Object 'ListView        'Un-AddRefed native class pointer for making Friend calls
    ThisPointer As Long
End Type

'========================================================================================
' API
'========================================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsEqualGUID Lib "ole32" (iid1 As Guid, iid2 As Guid) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Type Guid
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

'========================================================================================
' Constants and member variables
'========================================================================================

Private Const S_FALSE               As Long = 1
Private Const S_OK                  As Long = 0

Private IID_IOleInPlaceActiveObject As Guid
Private m_IPAOVTable(9)             As Long

Private Type tHookStruct
  hWnd As Long
  Struct As IPAOHookStructListView
End Type

Dim HookStructs(9999) As tHookStruct
Dim HookStructsCount As Long

'========================================================================================
' Functions
'========================================================================================

Function GetHookStructIndex(Ctl) As Long
  Dim i As Long
  For i = 0 To HookStructsCount - 1
    If HookStructs(i).hWnd = Ctl.hWnd Then
      GetHookStructIndex = i
      Exit For
    End If
  Next
End Function

Sub AddStruct(Ctl)
  'ReDim Preserve HookStructs(HookStructsCount)
  HookStructs(HookStructsCount).hWnd = Ctl.hWnd
  HookStructsCount = HookStructsCount + 1
End Sub

'***Public Sub InitIPAO(IPAOHookStruct As IPAOHookStructListView, Ctl As ListView)
Public Sub InitIPAO(Ctl As Object)
  Dim IPAO As IOleInPlaceActiveObject
  Call AddStruct(Ctl)
  With HookStructs(HookStructsCount - 1).Struct 'IPAOHookStruct
      Set IPAO = Ctl
      Call CopyMemory(.IPAOReal, IPAO, 4)
      Call CopyMemory(.Ctl, Ctl, 4)
      .lpVTable = GetVTable
      '***.ThisPointer = VarPtr(IPAOHookStruct)
      .ThisPointer = VarPtr(HookStructs(HookStructsCount - 1).Struct)
  End With
End Sub

'***Public Sub TerminateIPAO(IPAOHookStruct As IPAOHookStructListView)
Public Sub TerminateIPAO(Ctl)
  '***With IPAOHookStruct
  With HookStructs(GetHookStructIndex(Ctl)).Struct
      Call CopyMemory(.IPAOReal, 0&, 4)
      Call CopyMemory(.Ctl, 0&, 4)
  End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function GetVTable() As Long

    ' Set up the vTable for the interface and return a pointer to it
    If (m_IPAOVTable(0) = 0) Then
        m_IPAOVTable(0) = AddressOfFunction(AddressOf QueryInterface)
        m_IPAOVTable(1) = AddressOfFunction(AddressOf AddRef)
        m_IPAOVTable(2) = AddressOfFunction(AddressOf Release)
        m_IPAOVTable(3) = AddressOfFunction(AddressOf GetWindow)
        m_IPAOVTable(4) = AddressOfFunction(AddressOf ContextSensitiveHelp)
        m_IPAOVTable(5) = AddressOfFunction(AddressOf TranslateAccelerator)
        m_IPAOVTable(6) = AddressOfFunction(AddressOf OnFrameWindowActivate)
        m_IPAOVTable(7) = AddressOfFunction(AddressOf OnDocWindowActivate)
        m_IPAOVTable(8) = AddressOfFunction(AddressOf ResizeBorder)
        m_IPAOVTable(9) = AddressOfFunction(AddressOf EnableModeless)
        '--- init guid
        With IID_IOleInPlaceActiveObject
            .Data1 = &H117&
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
    End If
    GetVTable = VarPtr(m_IPAOVTable(0))
End Function

Private Function AddressOfFunction(lpfn As Long) As Long
    ' Work around, VB thinks lPtr = AddressOf Method is an error
    AddressOfFunction = lpfn
End Function

'========================================================================================
' Interface implemenattion
'========================================================================================

Private Function AddRef(This As IPAOHookStructListView) As Long
    AddRef = This.IPAOReal.AddRef
End Function

Private Function Release(This As IPAOHookStructListView) As Long
    Release = This.IPAOReal.Release
End Function

Private Function QueryInterface(This As IPAOHookStructListView, riid As Guid, pvObj As Long) As Long
    ' Install the interface if required
    If (IsEqualGUID(riid, IID_IOleInPlaceActiveObject)) Then
        ' Install alternative IOleInPlaceActiveObject interface implemented here
        pvObj = This.ThisPointer
        AddRef This
        QueryInterface = 0
      Else
        ' Use the default support for the interface:
        QueryInterface = This.IPAOReal.QueryInterface(ByVal VarPtr(riid), pvObj)
    End If
End Function

Private Function GetWindow(This As IPAOHookStructListView, phwnd As Long) As Long
    GetWindow = This.IPAOReal.GetWindow(phwnd)
End Function

Private Function ContextSensitiveHelp(This As IPAOHookStructListView, ByVal fEnterMode As Long) As Long
    ContextSensitiveHelp = This.IPAOReal.ContextSensitiveHelp(fEnterMode)
End Function

Private Function TranslateAccelerator(This As IPAOHookStructListView, lpMsg As msg) As Long
    ' Check if we want to override the handling of this key code:
    If (frTranslateAccel(lpMsg, This.Ctl)) Then
        TranslateAccelerator = S_OK
      Else
        ' If not pass it on to the standard UserControl TranslateAccelerator method:
        TranslateAccelerator = This.IPAOReal.TranslateAccelerator(ByVal VarPtr(lpMsg))
    End If
End Function

Private Function OnFrameWindowActivate(This As IPAOHookStructListView, ByVal fActivate As Long) As Long
    OnFrameWindowActivate = This.IPAOReal.OnFrameWindowActivate(fActivate)
End Function

Private Function OnDocWindowActivate(This As IPAOHookStructListView, ByVal fActivate As Long) As Long
    OnDocWindowActivate = This.IPAOReal.OnDocWindowActivate(fActivate)
End Function

Private Function ResizeBorder(This As IPAOHookStructListView, prcBorder As RECT, ByVal puiWindow As IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
    ResizeBorder = This.IPAOReal.ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
End Function

Private Function EnableModeless(This As IPAOHookStructListView, ByVal fEnable As Long) As Long
    EnableModeless = This.IPAOReal.EnableModeless(fEnable)
End Function

Public Function ShiftState() As Integer
  Dim lS As Integer
  If (GetAsyncKeyState(vbKeyShift) < 0) Then
      lS = lS Or vbShiftMask
  End If
  If (GetAsyncKeyState(vbKeyMenu) < 0) Then
      lS = lS Or vbAltMask
  End If
  If (GetAsyncKeyState(vbKeyControl) < 0) Then
      lS = lS Or vbCtrlMask
  End If
  ShiftState = lS
End Function

Public Sub pvSetIPAO(Ctl As Object)

  Dim pOleObject          As IOleObject
  Dim pOleInPlaceSite     As IOleInPlaceSite
  Dim pOleInPlaceFrame    As IOleInPlaceFrame
  Dim pOleInPlaceUIWindow As IOleInPlaceUIWindow
  Dim rcPos               As RECT
  Dim rcClip              As RECT
  Dim uFrameInfo          As OLEINPLACEFRAMEINFO

  On Error Resume Next
  Set pOleObject = Ctl
  Set pOleInPlaceSite = pOleObject.GetClientSite
  If (Not pOleInPlaceSite Is Nothing) Then
      Call pOleInPlaceSite.GetWindowContext(pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo))
      '***If (Not pOleInPlaceFrame Is Nothing) Then Call pOleInPlaceFrame.SetActiveObject(Ctl.uIPAO.ThisPointer, vbNullString)
      If (Not pOleInPlaceFrame Is Nothing) Then Call pOleInPlaceFrame.SetActiveObject(HookStructs(GetHookStructIndex(Ctl)).Struct.ThisPointer, vbNullString)


      If (Not pOleInPlaceUIWindow Is Nothing) Then '-- And Not m_bMouseActivate
          '***Call pOleInPlaceUIWindow.SetActiveObject(Ctl.uIPAO.ThisPointer, vbNullString)
          Call pOleInPlaceUIWindow.SetActiveObject(HookStructs(GetHookStructIndex(Ctl)).Struct.ThisPointer, vbNullString)
          
      Else
          Call pOleObject.DoVerb(OLEIVERB_UIACTIVATE, 0, pOleInPlaceSite, 0, Ctl.hWnd, VarPtr(rcPos))
      End If
  End If
  On Error GoTo 0

End Sub

Function frTranslateAccel(pMsg As msg, Ctl As Object) As Boolean

  Dim pOleObject As IOleObject
  Dim pOleControlSite As IOleControlSite
  Dim hEdit As Long

  On Error Resume Next

  Select Case pMsg.Message
  Case WM_KEYDOWN, WM_KEYUP

      Select Case pMsg.wParam
        Case vbKeyTab
          If (ShiftState() And vbCtrlMask) Then
              Set pOleObject = Ctl
              Set pOleControlSite = pOleObject.GetClientSite
              If (Not pOleControlSite Is Nothing) Then
                  Call pOleControlSite.TranslateAccelerator(VarPtr(pMsg), ShiftState() And vbShiftMask)
              End If
          End If
          frTranslateAccel = False
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
          hEdit = Ctl.pvEdithWnd()
          'fMain.Caption = hEdit 'pMsg.Message
          If (hEdit) Then
              Call SendMessageLong(hEdit, pMsg.Message, pMsg.wParam, pMsg.lParam)
          Else
              Call SendMessageLong(Ctl.hWnd, pMsg.Message, pMsg.wParam, pMsg.lParam)
          End If
          frTranslateAccel = True
      End Select

  End Select

  On Error GoTo 0

End Function
