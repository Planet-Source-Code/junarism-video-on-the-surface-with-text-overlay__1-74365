VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ResX = 800
Private Const ResY = 600

Private bRunning As Boolean
Private bCapture As Boolean

Private mHandle As Long

Private cDDraw As New clsDDraw
Private cDShow As New clsDShow
Private cWindow As New clsWindow

Private cText As New clsFXText
Private ifps As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function InitDirectDraw(ByVal hWnd As Long) As Boolean
    mHandle = hWnd
    If cDDraw.CreateDDraw(mHandle) Then
        cDDraw.CreateSurface 0, ResX, ResY
        cWindow.ResizeWindow mHandle, ResX, ResY
        InitDirectDraw = True
    End If
End Function

Public Sub SetFonts()
    cText.Init "Arial", 20, 70, &HFFFFFF, &H804000
End Sub

Public Sub SetCapture(ByVal filter As String, device As Boolean)
    StopCapture
    
    If device Then
        cDShow.DeviceList filter
        filter = Split(filter, vbTab)(0)
    End If
    
    If cDShow.CreateCapture(filter, Not device) Then
        Dim w As Long, h As Long
        cDShow.GetSize w, h
        cDDraw.CreateSurface 1, w, h
        cDShow.PlayCapture
        bCapture = True
    End If
End Sub

Public Sub StopCapture()
    bCapture = False
    Sleep 20
    cDShow.StopCapture
End Sub

Public Sub RenderScreen()
    Dim hdc As Long
    Dim rct As RECT
    Dim txt As String
        
    Do While Me.IsRunning
        If bCapture Then
            If cDShow.GetFrame Then
                cDDraw.GetSurfaceDC 1, hdc
                cDShow.SetFrame hdc
                cDDraw.SetSurfaceDC 1, hdc
            End If
            cDDraw.BlitSurface 0, 1
        Else
            cDDraw.FillSurface 0, 0
        End If

        cDDraw.GetSurfaceDC 0, hdc

        'txt = Now
        txt = "F1 - Video     F2 - Device     F3 - Stop"
        With cText
            .DrawText hdc, (ResX - .TextWidth(txt)) \ 2, 200, txt
        End With

        cDDraw.SetSurfaceDC 0, hdc

        cWindow.GetWindowRect mHandle, rct
        cDDraw.RenderSurface 0, rct
        
        ifps = ifps + 1
        
        DoEvents
        Sleep 10
    Loop
End Sub

Property Get IsRunning() As Boolean
    IsRunning = bRunning
End Property

Property Let IsRunning(value As Boolean)
    bRunning = value
    If value = False Then
        Sleep 20
    End If
End Property

Property Get FPS() As Integer
    FPS = ifps
End Property

Property Let FPS(value As Integer)
    ifps = value
End Property

Private Sub Class_Terminate()
    Set cWindow = Nothing
    Set cDShow = Nothing
    Set cDDraw = Nothing
End Sub
