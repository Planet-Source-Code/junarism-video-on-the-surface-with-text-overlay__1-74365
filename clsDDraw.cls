VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dd As DirectDraw2
Private ddClipper As DirectDrawClipper

Private ddsFront As DirectDrawSurface2
Private ddsdFront As DDSURFACEDESC

Private ddsBack(1) As DirectDrawSurface2
Private ddsdBack(1) As DDSURFACEDESC
Private rcBack(1) As RECT

Public Function CreateDDraw(ByVal hWnd As Long) As Boolean
    On Error GoTo Err
    
    DirectDrawCreate ByVal 0&, dd, Nothing
    dd.SetCooperativeLevel hWnd, DDSCL_NORMAL
    
    With ddsdFront
        .dwSize = Len(ddsdFront)
        .dwFlags = DDSD_CAPS
        .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE
    End With
    dd.CreateSurface ddsdFront, ddsFront, Nothing
    
    dd.CreateClipper 0, ddClipper, Nothing
    ddClipper.SetHWnd 0, hWnd
    ddsFront.SetClipper ddClipper
    
    CreateDDraw = True
    Exit Function
Err:
    Class_Terminate
End Function

Public Sub CreateSurface(ByVal idx As Integer, ByVal w As Long, ByVal h As Long)
    If Not (dd Is Nothing) Then
        Set ddsBack(idx) = Nothing
        With ddsdBack(idx)
            .dwSize = Len(ddsdBack(idx))
            .dwFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
            .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN
            .dwWidth = w
            .dwHeight = h
        End With
        With rcBack(idx)
            .Right = w
            .bottom = h
        End With
        dd.CreateSurface ddsdBack(idx), ddsBack(idx), Nothing
    End If
End Sub

Public Sub FillSurface(ByVal idx As Integer, ByVal color As Long)
    Dim fx As DDBLTFX
    With fx
        .dwSize = Len(fx)
        .dwFillColor = color
    End With
    ddsBack(idx).Blt rcBack(idx), Nothing, rcBack(idx), DDBLT_COLORFILL, fx
End Sub

Public Sub BlitSurface(ByVal idx1 As Integer, ByVal idx2 As Integer)
    On Error Resume Next
    Dim fx As DDBLTFX
    fx.dwSize = Len(fx)
    ddsBack(idx1).Blt rcBack(idx1), ddsBack(idx2), rcBack(idx2), DDBLT_WAIT, fx
End Sub

Public Sub GetSurfaceDC(ByVal idx As Integer, hdc As Long)
    ddsBack(idx).GetDC hdc
End Sub

Public Sub SetSurfaceDC(ByVal idx As Integer, hdc As Long)
    ddsBack(idx).ReleaseDC hdc
End Sub

Public Sub RenderSurface(ByVal idx As Integer, rcDest As RECT)
    On Error Resume Next
    Dim fx As DDBLTFX
    fx.dwSize = Len(fx)
    ddsFront.Blt rcDest, ddsBack(idx), rcBack(idx), DDBLT_WAIT, fx
End Sub

Private Sub Class_Terminate()
    If Not dd Is Nothing Then
        dd.SetCooperativeLevel 0, DDSCL_NORMAL
        Dim i As Integer
        For i = UBound(ddsBack) To 0 Step -1
            Set ddsBack(i) = Nothing
        Next
        Set ddsFront = Nothing
        Set ddClipper = Nothing
        Set dd = Nothing
    End If
End Sub
