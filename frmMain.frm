VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitBlt"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1650
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   1650
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   13000
      Left            =   -9480
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   12945
      ScaleWidth      =   11475
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   13000
      Left            =   2160
      Picture         =   "frmMain.frx":1DD562
      ScaleHeight     =   12945
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   11535
   End
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim KeyCheck As Integer
Private vUp, vDown, vLeft, vRight, vUpLeft, vUpRight, vDownLeft, vDownRight As Integer
Private Const cKeyUp = 1 ' c in the name means constant
Private Const cKeyDown = 2
Private Const cKeyLeft = 4
Private Const cKeyRight = 8
Private Const cUpLeft = 5
Private Const cUpRight = 9
Private Const cDownLeft = 6
Private Const cDownRight = 10

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TimerChecker As Boolean
    TimerChecker = False
    Select Case KeyCode
        Case vbKeyDown:
            KeyCheck = KeyCheck Or cKeyDown
            TimerChecker = True
        Case vbKeyUp:
            KeyCheck = KeyCheck Or cKeyUp
            TimerChecker = True
        Case vbKeyLeft
            KeyCheck = KeyCheck Or cKeyLeft
            TimerChecker = True
        Case vbKeyRight
            KeyCheck = KeyCheck Or cKeyRight
            TimerChecker = True
    End Select
    If TimerChecker Then
        If Not Timer1.Enabled Then
            Call Timer1_Timer
            Timer1.Enabled = True
        End If
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown:
            KeyCheck = KeyCheck And (Not cKeyDown)
        Case vbKeyUp:
            KeyCheck = KeyCheck And (Not cKeyUp)
        Case vbKeyLeft
            KeyCheck = KeyCheck And (Not cKeyLeft)
        Case vbKeyRight
            KeyCheck = KeyCheck And (Not cKeyRight)
    End Select
    If KeyCheck = 0 Then Timer1.Enabled = False
End Sub

Private Sub Form_Load()
MsgBox ("Welcome to Step3 of my Game Project:Blackout Forest.Use the Arrow Keys to control the character!(Dont forget that u can move him by presing eg.Down & Right Arrow),Have fun!")
Call Mover(1, 761)
End Sub

Private Sub Timer1_Timer()
    If KeyCheck = 0 Then Exit Sub
    If (KeyCheck And cUpLeft) = cUpLeft Then
        vUpLeft = vUpLeft + 1
        Call Mover(vUpLeft, 286)
        vUpLeft = vUpLeft + 94
        If vUpLeft > 94 * 8 + 1 Then vUpLeft = 0
ElseIf (KeyCheck And cUpRight) = cUpRight Then
        vUpRight = vUpRight + 1
        Call Mover(vUpRight, 476)
        vUpRight = vUpRight + 94
        If vUpRight > 94 * 8 + 1 Then vUpRight = 0
    ElseIf (KeyCheck And cDownLeft) = cDownLeft Then
        vDownLeft = vDownLeft + 1
        Call Mover(vDownLeft, 96)
        vDownLeft = vDownLeft + 94
        If vDownLeft > 94 * 8 + 1 Then vDownLeft = 0
    ElseIf (KeyCheck And cDownRight) = cDownRight Then
        vDownRight = vDownRight + 1
        Call Mover(vDownRight, 666)
        vDownRight = vDownRight + 94
        If vDownRight > 94 * 8 + 1 Then vDownRight = 0
    ElseIf (KeyCheck And cKeyUp) = cKeyUp Then
        vUp = vUp + 1
        Call Mover(vUp, 381)
        vUp = vUp + 94
        If vUp > 94 * 8 + 1 Then vUp = 0
    ElseIf (KeyCheck And cKeyDown) = cKeyDown Then
        vDown = vDown + 1
        Call Mover(vDown, 1)
        vDown = vDown + 94
        If vDown > 94 * 8 + 1 Then vDown = 0
    ElseIf (KeyCheck And cKeyLeft) = cKeyLeft Then
            vLeft = vLeft + 1
        Call Mover(vLeft, 191)
        vLeft = vLeft + 94
        If vLeft > 94 * 8 + 1 Then vLeft = 0
    ElseIf (KeyCheck And cKeyRight) = cKeyRight Then
            vRight = vRight + 1
        Call Mover(vRight, 571)
        vRight = vRight + 94
        If vRight > 94 * 8 + 1 Then vRight = 0
    End If
End Sub

Function Mover(xSrc, ySrc)
Me.Cls
Call BitBlt(Me.hDC, 0, 0, 94, 94, Picture1.hDC, xSrc, ySrc, vbSrcAnd)
Call BitBlt(Me.hDC, 0, 0, 94, 94, Picture2.hDC, xSrc, ySrc, vbSrcInvert)
End Function
