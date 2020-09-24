VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Settings"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   5670
   Begin VB.Frame fraList 
      BackColor       =   &H00800000&
      Caption         =   "List Placement"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton optRight 
         BackColor       =   &H00800000&
         Caption         =   "Align List to Right of Window"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton optLeft 
         BackColor       =   &H00800000&
         Caption         =   "Align List to Left of Window"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame fraToolbar 
      BackColor       =   &H00800000&
      Caption         =   "Toolbar Icons"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton optSmall 
         BackColor       =   &H00800000&
         Caption         =   "Use Small Toolbar Icons"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optLarge 
         BackColor       =   &H00800000&
         Caption         =   "Use Large Toolbar Icons"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Transparency"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
      Begin VB.TextBox txtTransDeg 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "175"
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton optTransOn 
         BackColor       =   &H00800000&
         Caption         =   "Turn On Transparency"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optTransOff 
         BackColor       =   &H00800000&
         Caption         =   "Turn Off Transparency"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblTransDeg 
         BackStyle       =   0  'Transparent
         Caption         =   "Degree of Transparency"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.TextBox txtDBPath 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Shows current location of database."
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Settings"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      ToolTipText     =   "Save/Apply settings."
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Cancel/Discard changes."
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblDBPath 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Path to the Current Database File"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '============================================================
    '
    '============================================================

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim lngVersion As Long, IconSize, AlignList
    Me.Icon = LoadResPicture(111, 1)
    txtDBPath.Text = DBPath
    lngVersion = GetVersion()
    AlignList = GetSetting("SourceCode", "Settings", "AlignList", 3)
    Select Case AlignList
        Case 3
            optLeft.Value = True
            optRight.Value = False
        Case 4
            optLeft.Value = False
            optRight.Value = True
    End Select
    If lngVersion = 143851525 Then
        TransReset
    Else
        optTransOn.Enabled = False
        optTransOff.Enabled = False
    End If
End Sub

Private Sub optTransOff_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    txtTransDeg.Text = GetSetting("SourceCode", "Settings", "TransDeg", 175)
    txtTransDeg.Enabled = False
    lblTransDeg.Enabled = False
    txtTransDeg.BackColor = &H8000000F
End Sub

Private Sub optTransOn_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    txtTransDeg.Enabled = True
    lblTransDeg.Enabled = True
    txtTransDeg.BackColor = &H80000005
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim DegTrans, AlignList
    Dim lngVersion As Long
    lngVersion = GetVersion()
    DegTrans = txtTransDeg.Text
    If lngVersion = 143851525 Then
        If DegTrans >= 0 And DegTrans <= 255 Then
            DegTrans = DegTrans
        ElseIf DegTrans > 255 Then
            DegTrans = 255
        ElseIf DegTrans < 0 Then
            DegTrans = 0
        Else
            DegTrans = InputBox("Please enter an integer between 0 and 255", "Bad Input Detected", 175)
        End If
        If optTransOff.Value = True Then
            SetLayered frmMain.hWnd, False, 255
            SaveSetting "SourceCode", "Settings", "TransState", "0"
        ElseIf optTransOn.Value = True Then
            SetLayered frmMain.hWnd, True, Int(DegTrans)
            SaveSetting "SourceCode", "Settings", "TransState", "1"
        End If
    End If
    If optLeft.Value = True And optRight.Value = False Then
        AlignList = 3
    ElseIf optRight.Value = True And optLeft.Value = False Then
        AlignList = 4
    End If
    SaveSetting "SourceCode", "Settings", "TransDeg", DegTrans
    SaveSetting "SourceCode", "Settings", "AlignList", AlignList
    frmMain.picList.Align = AlignList
    Unload Me
    frmMain.MDIForm_Load
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    TransReset
    Unload Me
End Sub

Private Sub TransReset()
On Error Resume Next
    '============================================================
    '
    '============================================================
    optTransOff_Click
    optTransOff.Value = True
    optTransOn.Value = False
End Sub

Private Sub optLarge_Click()
    '============================================================
    '
    '============================================================
    optSmall.Value = False
    optLarge.Value = True
    SaveSetting "SourceCode", "Settings", "Icons", 0
End Sub

Private Sub optSmall_Click()
    '============================================================
    '
    '============================================================
    optSmall.Value = True
    optLarge.Value = False
    SaveSetting "SourceCode", "Settings", "Icons", 1
End Sub
