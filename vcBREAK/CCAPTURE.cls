VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCAPTURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------
' Very old reusable class that allows you to capture
' the screen or objects by their hWnd.
'------------------------------------------------------------

'------------------------------------------------------------
' It has been a long time since I used this [before
' ME and XP systems] so not 100 percent sure it
' will work on all OSes anymore.
'------------------------------------------------------------

Option Explicit
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Public FileName As String
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Sub ScrnCap(Lt, Top, Rt, Bot)
    On Error GoTo ErrorScrnCap
    Dim rWIDTH As Long, rHEIGHT As Long
    Dim SourceDC As Long, DestDC As Long, bHANDLE As Long, Wnd As Long
    Dim dHANDLE As Long, dm As DEVMODE
    rWIDTH = Rt - Lt
    rHEIGHT = Bot - Top
    SourceDC = CreateDC("DISPLAY", 0&, 0&, dm)
    DestDC = CreateCompatibleDC(SourceDC)
    bHANDLE = CreateCompatibleBitmap(SourceDC, rWIDTH, rHEIGHT)
    SelectObject DestDC, bHANDLE
    BitBlt DestDC, 0, 0, rWIDTH, rHEIGHT, SourceDC, Lt, Top, &HCC0020
    Wnd = 0
    OpenClipboard Wnd
    EmptyClipboard
    SetClipboardData 2, bHANDLE
    CloseClipboard
    DeleteDC DestDC
    ReleaseDC dHANDLE, SourceDC
    Exit Sub
ErrorScrnCap:
    MsgBox Err & ":Error in ScrnCap().  Error Message:" & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@hotmail.net]
' Purpose:  Captures a control/window based on hWnd and save to a .BMP.
' Parameters:  hWnd=Control/Window to capture, fNAME=Save .BMP to...
' Example:  obj.Capture lvLIST.hWnd,"TEST.BMP"
' Date: August,21 1998 @ 16:18:27
'------------------------------------------------------------
Public Sub Capture(control_hWnd As Long, Optional fNAME As String = "", Optional OnlyToClipBoard As Boolean = False)
    On Error GoTo ErrorCapture
    Dim sp As RECT, x As Long
    If fNAME = "" Then
        fNAME = Me.FileName
    Else
        Me.FileName = fNAME
    End If
    If Me.FileName <> "" Then
        x = GetWindowRect(control_hWnd, sp)
        ScrnCap sp.Left, sp.Top, sp.Right, sp.Bottom
        If OnlyToClipBoard = False Then
            SavePicture Clipboard.GetData, Me.FileName
        End If
    End If
    Exit Sub
ErrorCapture:
    MsgBox Err & ":Error in Caputre().  Error Message:" & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Public Sub CaptureDesktop()
    On Error GoTo ErrorCaptureDesktop
    Dim dhWND As Long, sp As RECT, x As Long
    dhWND = GetDesktopWindow
    If dhWND <> 0 Then
        x = GetWindowRect(dhWND, sp)
        ScrnCap sp.Left, sp.Top, sp.Right, sp.Bottom
    End If
    Exit Sub
ErrorCaptureDesktop:
    MsgBox Err & ":Error in CaptureDesktop.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub


