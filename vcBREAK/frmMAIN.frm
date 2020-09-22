VERSION 5.00
Begin VB.Form frmMAIN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "vcBreak"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' Simply little program written just because.
'------------------------------------------------------------
' For more software developed by me [much better
' than this] feel free to visit my site at:
'
' http://vbasic.iscool.net
'
'------------------------------------------------------------
Option Explicit
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        '------------------------------------------------------------
        ' Right Click to Exit
        '------------------------------------------------------------
        If MsgBox("Click yes to close, no to reset.", vbQuestion + vbYesNo) = vbYes Then
            Unload Me
        Else
            Me.Cls
        End If
    Else
        '------------------------------------------------------------
        ' else break at point clicked
        '------------------------------------------------------------
        BreakIt x, y
    End If
End Sub
Private Sub Form_Load()
    Dim obj As CCAPTURE
    Set obj = New CCAPTURE
    '------------------------------------------------------------
    ' Capture the screen and assign it to the picture
    ' property of the form.
    '------------------------------------------------------------
    obj.CaptureDesktop
    Set Me.Picture = Clipboard.GetData
    '------------------------------------------------------------
    ' maximize the form [forms border is set to none
    ' so it will take up the whole screen]
    '------------------------------------------------------------
    Me.WindowState = vbMaximized
    '------------------------------------------------------------
    ' Set this for later when I print some text
    '------------------------------------------------------------
    Me.FontSize = 12
    Me.FontBold = True
    Me.DrawWidth = 1
    Me.FillStyle = 0
    Me.FillColor = vbBlack
End Sub
Public Sub BreakIt(Optional x As Single = -1, Optional y As Single = -1)
    On Error GoTo ErrorBreakIt
    Dim cx As Long, cy As Long
    Dim nx As Long, ny As Long
    Dim ax As Long, ay As Long
    Dim lCNT As Long, lSize As Long, lMsg As String
    Dim sx As Long, sy As Long
    Randomize
    With Me
        '------------------------------------------------------------
        ' Find center or use point supplied
        '------------------------------------------------------------
        If x <> -1 And y <> -1 Then
            cx = x
            cy = y
        Else
            cx = .Width / 2
            cy = .Height / 2
        End If
        '------------------------------------------------------------
        ' Just drawing a small circle at center point trying
        ' to get an effect.
        '------------------------------------------------------------
        Me.Circle (cx, cy), 20, vbBlack
        '------------------------------------------------------------
        ' Display instructions to user at the top of the
        ' screen centered
        '------------------------------------------------------------
        lMsg = "Left click to break, right click to exit."
        Me.CurrentY = 0
        Me.CurrentX = (Screen.Width / 2) - (Me.TextWidth(lMsg) / 2)
        Me.Print lMsg
        '------------------------------------------------------------
        ' Random size of how big to make the lines
        '------------------------------------------------------------
        lSize = Int(Rnd * 600) + 200
        '------------------------------------------------------------
        ' Loop random number of times (10-20 times)
        '------------------------------------------------------------
        For lCNT = 1 To Int(Rnd * 20) + 10
            '------------------------------------------------------------
            ' Pick a random point
            '------------------------------------------------------------
            nx = Int(Rnd * lSize) + 1
            ny = Int(Rnd * lSize) + 1
            If nx Mod 2 = 0 Then nx = 0 - nx
            If ny Mod 2 = 0 Then ny = 0 - ny
            '------------------------------------------------------------
            ' Draw a black line from center (or given point)  to random point
            '------------------------------------------------------------
            Me.Line (cx, cy)-(cx - nx, cy - ny), vbBlack
            If (Int(Rnd * 100) + 1) Mod 2 = 0 Then
                sx = Int(Rnd * 500) + 100
                sy = Int(Rnd * 500) + 100
                If nx > 0 And ny > 0 Then
                    Me.Line (cx - nx, cy - ny)-(cx - nx - sx, cy - ny - sy), vbBlack
                End If
                If nx > 0 And ny < 0 Then
                    Me.Line (cx - nx, cy - ny)-(cx - nx - sx, cy - ny + sy), vbBlack
                End If
                If nx < 0 And ny < 0 Then
                    Me.Line (cx - nx, cy - ny)-(cx - nx + sx, cy - ny + sy), vbBlack
                End If
                If nx < 0 And ny > 0 Then
                    Me.Line (cx - nx, cy - ny)-(cx - nx + sx, cy - ny - sy), vbBlack
                End If
            End If
        Next
    End With
    '------------------------------------------------------------
    ' Play wav file from resource file of glass breaking.
    '------------------------------------------------------------
    PlayWaveRes sndGLASS
    Exit Sub
ErrorBreakIt:
    MsgBox Err & ":Error in call to BreakIt()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
