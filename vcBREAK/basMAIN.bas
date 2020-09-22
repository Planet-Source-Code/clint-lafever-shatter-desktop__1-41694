Attribute VB_Name = "basMAIN"
Option Explicit
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SND_ASYNC = &H1
Private bytSound() As Byte
Public Enum SoundFlags
    soundSYNC = SND_SYNC
    soundNO_DEFAULT = SND_NODEFAULT
    soundMEMORY = SND_MEMORY
    soundLOOP = SND_LOOP
    soundNO_STOP = SND_NOSTOP
    soundASYNC = SND_ASYNC
End Enum
Public Enum AppSounds
    sndGLASS = 101
End Enum
'------------------------------------------------------------
' This sub routine plus the API and contants above
' are used to access a WAV file stored in the Custom
' section of a resource file.  It can then play
' the WAV file
'------------------------------------------------------------
Public Sub PlayWaveRes(vntResourceID As AppSounds, Optional vntFlags As SoundFlags = soundASYNC)
    bytSound = LoadResData(vntResourceID, "WAVE")
    If IsMissing(vntFlags) Then
        vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
    End If
    If (vntFlags And SND_MEMORY) = 0 Then
        vntFlags = vntFlags Or SND_MEMORY
    End If
    sndPlaySound bytSound(0), vntFlags
End Sub

'------------------------------------------------------------
' Start the program off
'------------------------------------------------------------
Public Sub Main()
    On Error GoTo ErrorMain
    Dim frm As frmMAIN
    '------------------------------------------------------------
    ' Open frmMAIN
    '------------------------------------------------------------
    Set frm = New frmMAIN
    frm.Show
    '------------------------------------------------------------
    ' Make sure it is ready
    '------------------------------------------------------------
    frm.Refresh: DoEvents
    '------------------------------------------------------------
    ' Call the BreakIt sub on that form
    '------------------------------------------------------------
    frm.BreakIt
    Exit Sub
ErrorMain:
    MsgBox Err & ":Error in call to Main()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
