VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "About AboutForm"
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      ToolTipText     =   "Humana Inc."
      Top             =   120
      Width           =   540
   End
   Begin VB.TextBox txtCopyright 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "frmAbout.frx":0614
      ToolTipText     =   "Author/Copyright information"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "OK"
      ToolTipText     =   "Close this form"
      Top             =   2040
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Tag             =   "&System Info..."
      ToolTipText     =   "Get system information"
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label lblComp 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   1050
      TabIndex        =   3
      Tag             =   "App Description"
      ToolTipText     =   "Program Description"
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1050
      TabIndex        =   2
      Tag             =   "Application Title"
      ToolTipText     =   "Program Title"
      Top             =   240
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225
      X2              =   5657
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================
' This is the standard about form.  I am not going to bother
' to comment this.  It can be generated automatically.
'============================================================

Const KEY_ALL_ACCESS = &H2003F
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
On Error Resume Next
    Me.Icon = LoadResPicture(120, 1)
    frmAbout.Caption = App.Title & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = App.Comments
    lblComp.Caption = "Compatible on " & App.FileDescription & " systems."
    DoGradient Me
End Sub

Private Sub cmdSysInfo_Click()
On Error Resume Next
    Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
        Unload Me
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
        Dim rc As Long
        Dim SysInfoPath As String
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                Else
                        GoTo SysInfoErr
                End If
        Else
                GoTo SysInfoErr
        End If
        Call Shell(SysInfoPath, vbNormalFocus)
        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long
        Dim rc As Long
        Dim hKey As Long
        Dim hDepth As Long
        Dim KeyValType As Long
        Dim tmpVal As String
        Dim KeyValSize As Long
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
        tmpVal = String$(1024, 0)
        KeyValSize = 1024
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        Select Case KeyValType
        Case REG_SZ
                KeyVal = tmpVal
        Case REG_DWORD
                For i = Len(tmpVal) To 1 Step -1
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
                Next
                KeyVal = Format$("&h" + KeyVal)
        End Select
        GetKeyValue = True
        rc = RegCloseKey(hKey)
        Exit Function
GetKeyError:
        KeyVal = ""
        GetKeyValue = False
        rc = RegCloseKey(hKey)
End Function
