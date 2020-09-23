VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFXText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const PROOF_QUALITY = 2
Private Const PS_SOLID = 0

Private Type TextMetric
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type POINTAPI
    cx As Long
    cy As Long
End Type

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TextMetric) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function CreateFontA Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Sub DeleteDC Lib "gdi32" (ByVal hdc As Long)
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private mHdc As Long
Private mFont As Long
Private mPen As Long
Private mBrush As Long

Sub Init(ByVal fnt As String, ByVal w As Long, ByVal h As Long, ByVal sc As Long, ByVal fc As Long)
    mFont = CreateFontA(h, w, 0, 0, FW_BOLD, 0, 0, 0, 0, 0, 0, PROOF_QUALITY, 0, fnt)
    mPen = CreatePen(PS_SOLID, 1, sc)
    mBrush = CreateSolidBrush(fc)
    mHdc = CreateCompatibleDC(GetDC(0))
    SelectObject mHdc, mFont
End Sub

Sub DrawText(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal txt As String)
    SelectObject hdc, mFont
    SelectObject hdc, mPen
    SelectObject hdc, mBrush
        
    SetBkMode hdc, TRANSPARENT
    BeginPath hdc
    TextOut hdc, x, y, txt, Len(txt)
    EndPath hdc
    StrokeAndFillPath hdc
End Sub

Property Get TextWidth(ByVal txt As String) As Long
    With GetTextSize(mHdc, txt)
        TextWidth = .cx
    End With
End Property

Property Get TextHeight(ByVal txt As String) As Long
    With GetTextSize(mHdc, txt)
        TextHeight = .cy
    End With
End Property

Property Get FontWidth() As Long
    FontWidth = GetFontWidth(mHdc)
End Property

Private Function GetTextSize(hdc As Long, txt As String) As POINTAPI
    Dim sz As POINTAPI
    Call GetTextExtentPoint32(hdc, txt, Len(txt), sz)
    With GetTextSize
        .cx = sz.cx + 1
        .cy = sz.cy + 1
    End With
End Function

Private Function GetFontWidth(hdc As Long) As Long
    Dim tm As TextMetric
    GetTextMetrics hdc, tm
    GetFontWidth = CInt(tm.tmAveCharWidth)
End Function

Private Sub Class_Terminate()
    DeleteObject mFont
    DeleteObject mPen
    DeleteObject mBrush
    DeleteDC mHdc
End Sub
