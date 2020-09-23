VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Sample"
   ClientHeight    =   4650
   ClientLeft      =   2820
   ClientTop       =   630
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3180
      Top             =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cMain As New clsMAIN

Private Sub Form_Load()
    If cMain.InitDirectDraw(Me.hWnd) Then
        cMain.SetFonts
        cMain.IsRunning = True
        Me.Show
        cMain.RenderScreen
    Else
        MsgBox "DirectDraw Cannot Initialized!", vbCritical, App.Title
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF1
            cMain.SetCapture App.Path & "\video.mpg", False 'Video
        Case vbKeyF2
            cMain.SetCapture "", True 'Device
        Case vbKeyF3
            cMain.StopCapture
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cMain.IsRunning = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cMain = Nothing
End Sub

Private Sub Timer1_Timer()
    Me.Caption = cMain.FPS
    cMain.FPS = 0
End Sub
