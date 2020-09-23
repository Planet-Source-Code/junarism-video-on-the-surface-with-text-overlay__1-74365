VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moManager As IAMMediaControl
Private moGrabber As IAMSampleGrabber
Private moPosition As IAMMediaPosition
Private muBitmap As TGDIBitmapInfoHeader
Private bf() As Long
    
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As TGDIBitmapInfoHeader, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As TGDIBitmapInfoHeader, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Function CreateCapture(ByVal fltname As String, Optional ByVal srcfile As Boolean = True) As Boolean
    On Error GoTo err
    
    Dim oRegFilter As IAMRegFilterInfo
    Dim tMediaType As TAMMediaType
    Dim oFltGrabber As IAMFilterInfo
    Dim oSmpGrabber As IAMSampleGrabber
    Dim oFltCapture As IAMFilterInfo
    Dim oPinOut As IAMPinInfo
    Dim oPinIn As IAMPinInfo
    
    Call StopCapture
    
    Set moManager = New AMFiltergraphManager
    
    For Each oRegFilter In moManager.RegFilterCollection
        If (oRegFilter.Name = amFilterNameGrabber) Then
            With tMediaType
                .MajorType = UUIDFromString(amIDMediaTypeVideo)
                .SubType = UUIDFromString(amIDMediaTypeVideoRGB24)
            End With
            Set oFltGrabber = oRegFilter.filter()
            Set oSmpGrabber = oFltGrabber.filter
            With oSmpGrabber
                .MediaType = tMediaType
                .SetBufferSamples 1&
                .SetOneShot 0&
            End With
        End If
    Next
    
    If Not (oFltGrabber Is Nothing) Then
        If srcfile Then
            Set oFltCapture = moManager.AddSourceFilter(fltname)
        Else
            For Each oRegFilter In moManager.RegFilterCollection
                If (oRegFilter.Name = fltname) Then
                    Set oFltCapture = oRegFilter.filter
                    Exit For
                End If
            Next
        End If
        
        If Not (oFltCapture Is Nothing) Then
            For Each oPinOut In oFltCapture.Pins
                If oPinOut.Direction = 1 Then
                    Exit For
                End If
            Next
            
            For Each oPinIn In oFltGrabber.Pins
                If oPinIn.Direction = 0 Then
                    oPinOut.Connect oPinIn
                    Exit For
                End If
            Next
            
            For Each oPinOut In oFltGrabber.Pins
                If oPinOut.Direction = 1 Then
                    oPinOut.Render
                    Exit For
                End If
            Next
        
            Set moGrabber = oSmpGrabber
            CreateCapture = True
            Exit Function
        End If
    End If
    
err:
    Set moManager = Nothing
End Function

Public Sub PlayCapture()
    If Not (moManager Is Nothing) Then
        Dim oVideoWindow As IAMVideoWindow
        Set oVideoWindow = moManager
        oVideoWindow.AutoShow = 0&
        Set moPosition = moManager
        moManager.Run
    End If
End Sub

Public Sub StopCapture()
    If Not (moGrabber Is Nothing) Then
        moGrabber.SetBufferSamples 0&
        Set moGrabber = Nothing
    End If
        
    If Not (moManager Is Nothing) Then
        Call moManager.Stop
        Set moManager = Nothing
    End If
End Sub

Public Sub GetSize(w As Long, h As Long)
    If Not (moManager Is Nothing) Then
        Dim oBasicVideo As IAMBasicVideo
        Set oBasicVideo = moManager
        oBasicVideo.GetVideoSize w, h
        With muBitmap
            .HeaderSize = Len(muBitmap)
            .Planes = 1
            .BitCount = 24
            .Width = w
            .Height = h
        End With
    End If
End Sub

Public Function GetFrame() As Boolean
    On Error GoTo err
    Dim bs As Long
    Call moGrabber.GetCurrentBuffer(bs, ByVal 0&)
    ReDim bf(bs \ 4 - 1)
    Call moGrabber.GetCurrentBuffer(bs, bf(0))
    GetFrame = True
err:
End Function

Public Sub SetFrame(hdc As Long)
    SetDIBitsToDevice hdc, 0, 0, muBitmap.Width, muBitmap.Height, 0, 0, 0, muBitmap.Height, bf(0), muBitmap, 0
End Sub

Public Function IsPlaying() As Boolean
    On Error GoTo err
    If Not (moManager Is Nothing) Then
        Dim oMediaEvent As IAMMediaEvent
        Set oMediaEvent = moManager
        oMediaEvent.WaitForCompletion 0
    End If
err:
    If err.Number = 287 Then
        IsPlaying = True
    End If
End Function

Public Sub DeviceList(dev As String)
    Dim oDevEnum As IAMCreateDevEnum
    Dim oEnum As ICoEnumMoniker
    Dim oMoniker As ICoMoniker
    Dim oPropBag As ICoPropertyBag
    Dim oPtr As Long
    Dim vName As Variant
    Dim iRet As Long
    
    Set oDevEnum = New AMSystemDeviceEnumerator
    'CLSID_VideoInputDeviceCategory
    oDevEnum.CreateClassEnumerator UUIDFromString("{860BB310-5D01-11d0-BD3B-00A0C911CE86}"), oEnum, 0&
    iRet = oEnum.Next(1&, oMoniker)
    Do While iRet = 1
        'IPropertyBag
        oPtr = ObjPtr(oMoniker.BindToStorage(Nothing, Nothing, UUIDFromString("{55272A00-42CB-11CE-8135-00AA004BB851}")))
        CopyMemory oPropBag, oPtr, 4
        oPropBag.Read "FriendlyName", vName, Nothing
        CopyMemory oPropBag, 0, 4
        dev = dev & vName & vbTab
        Set oMoniker = Nothing
        iRet = oEnum.Next(1&, oMoniker)
    Loop
    
    Set oEnum = Nothing
    Set oDevEnum = Nothing
End Sub

Private Sub Class_Terminate()
    StopCapture
End Sub
