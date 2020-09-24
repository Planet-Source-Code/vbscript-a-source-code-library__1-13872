Attribute VB_Name = "modUtilities"
Option Explicit
'============================================================
'
'============================================================
Public Const mlngWindows95 = 0
Public Const mlngWindowsNT = 1
Public Const mlngWindows2000 = 2
Public Declare Function GetVersion Lib "kernel32" () As Long
Public glngWhichWindows32 As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Public Type POINTAPI
    x As Long
    y As Long
    End Type

Public Type SIZE
    cx As Long
    cy As Long
    End Type

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
    End Type
    Public Const WS_EX_LAYERED = &H80000
    Public Const GWL_STYLE = (-16)
    Public Const GWL_EXSTYLE = (-20)
    Public Const AC_SRC_OVER = &H0
    Public Const AC_SRC_ALPHA = &H1
    Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
    Public Const AC_SRC_NO_ALPHA = &H2
    Public Const AC_DST_NO_PREMULT_ALPHA = &H10
    Public Const AC_DST_NO_ALPHA = &H20
    Public Const LWA_COLORKEY = &H1
    Public Const LWA_ALPHA = &H2
    Public Const ULW_COLORKEY = &H1
    Public Const ULW_ALPHA = &H2
    Public Const ULW_OPAQUE = &H4
    Public lret As Long

Function CheckLayered(ByVal hWnd As Long) As Boolean
On Error Resume Next
    '============================================================
    '
    '============================================================
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (lret And WS_EX_LAYERED) = WS_EX_LAYERED Then
        CheckLayered = True
    Else
        CheckLayered = False
    End If
End Function

Function SetLayered(ByVal hWnd As Long, SetAs As Boolean, bAlpha As Byte)
On Error Resume Next
    '============================================================
    '
    '============================================================
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)
    If SetAs = True Then
        lret = lret Or WS_EX_LAYERED
    Else
        lret = lret And Not WS_EX_LAYERED
    End If
    SetWindowLong hWnd, GWL_EXSTYLE, lret
    SetLayeredWindowAttributes hWnd, 0, bAlpha, LWA_ALPHA
End Function

Public Sub DoTrans(FormName As Object)
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim lngVersion As Long
    lngVersion = GetVersion()
    If lngVersion = 143851525 Then
        SetLayered FormName.hWnd, True, 150
    End If
End Sub

Public Sub DoGradient(FormName As Object)
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim i As Integer, y As Integer
    FormName.AutoRedraw = True
    FormName.DrawStyle = 6
    FormName.DrawMode = 13
    FormName.DrawWidth = 13
    FormName.ScaleMode = 3
    FormName.ScaleHeight = 256
    For i = 0 To 510
        FormName.Line (0, y)-(FormName.Width, y + 1), RGB(0, 0, i), BF
        y = y + 1
    Next i
End Sub

Public Function DBPath() As String
On Error Resume Next
    '============================================================
    '
    '============================================================
    DBPath = GetSetting("SourceCode", "Settings", "Path")
    If DBPath = "" Then
        DBPath = InputBox("Please type the full path to your database.", "Database Not Found", App.Path & "\codelib.mdb")
        SaveSetting "SourceCode", "Settings", "Path", DBPath
    End If
End Function

Public Sub Main()
    '============================================================
    '
    '============================================================
    Dim RunOnce As String, DateRan As String
    RunOnce = GetSetting("SourceCode", "RunOnce", "RunOnce")
    DateRan = Date
    Select Case RunOnce
        Case 1
            frmMain.Show
        Case Else
            SaveSetting "SourceCode", "RunOnce", "RunOnce", "1"
            SaveSetting "SourceCode", "RunOnce", "DateRan", DateRan
            SaveSetting "SourceCode", "Settings", "AlignList", "3"
            SaveSetting "SourceCode", "Settings", "DataSet", "Source"
            SaveSetting "SourceCode", "Settings", "Icons", "0"
            SaveSetting "SourceCode", "Settings", "Path", ""
            SaveSetting "SourceCode", "Settings", "TransDeg", "128"
            SaveSetting "SourceCode", "Settings", "TransState", "0"
            frmMain.Show
    End Select
End Sub
