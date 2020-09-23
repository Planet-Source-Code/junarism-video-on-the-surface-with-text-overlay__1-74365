VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const WS_THICKFRAME = &H40000
Private Const WS_CAPTION = &HC00000  'WS_BORDER Or WS_DLGFRAME
Private Const WS_SYSMENU = &H80000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000

Private Const GWL_STYLE = (-16)

Private Declare Function AdjustWindowRect Lib "user32" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRectA Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Sub ResizeWindow(ByVal hWnd As Long, ByVal w As Long, ByVal h As Long)
    Dim r As RECT
    r.Right = w
    r.bottom = h
    AdjustWindowRect r, WS_THICKFRAME Or WS_CAPTION, False
    MoveWindow hWnd, 0, 0, r.Right - r.Left, r.bottom - r.Top, 1
End Sub

Sub GetWindowRect(ByVal hWnd As Long, r As RECT)
    Dim p As POINTAPI
    ClientToScreen hWnd, p
    GetClientRect hWnd, r
    OffsetRect r, p.x, p.y
End Sub

Sub SetWindowState(ByVal hWnd As Long)
    Static IsFullScreen As Boolean
    Static OldStyle As Long
    Static OldRECT As RECT
    
    IsFullScreen = Not IsFullScreen
    
    If IsFullScreen Then
        Dim hStyle As Long
        Dim r As RECT
        
        OldStyle = GetWindowLong(hWnd, GWL_STYLE)
        GetWindowRectA hWnd, OldRECT
        
        hStyle = OldStyle
        hStyle = hStyle And Not WS_THICKFRAME
        hStyle = hStyle And Not WS_CAPTION
        hStyle = hStyle And Not WS_SYSMENU
        hStyle = hStyle And Not WS_MAXIMIZEBOX
        hStyle = hStyle And Not WS_MINIMIZEBOX
        SetWindowLong hWnd, GWL_STYLE, hStyle
        
        GetWindowRectA GetDesktopWindow, r
        MoveWindow hWnd, 0, 0, r.Right - r.Left, r.bottom - r.Top, 1
    Else
        If OldStyle <> 0 Then
            SetWindowLong hWnd, GWL_STYLE, OldStyle
            With OldRECT
                MoveWindow hWnd, .Left, .Top, .Right - .Left, .bottom - .Top, 1
            End With
        End If
    End If
End Sub
