VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProjView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Code"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   8505
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Send Code to Clipboard"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      ToolTipText     =   "Copies full text from the Source Code window into the clipboard to be pasted in your application."
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtPhone 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtCodeName 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtRelDate 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtPurpose 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtAuthor 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   12648447
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmProjView.frx":0000
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Author"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Phone/Extension"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblCodeName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code Name/Title"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblOrigDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Release Date"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblPurpose 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purpose/Function of the Project"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Source Code - Cut and paste into your document"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   8295
   End
End
Attribute VB_Name = "frmProjView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '============================================================
    '
    '============================================================
    Dim db As Database
    Dim rsCode As Recordset

Private Sub cmdExit_Click()
    '============================================================
    '
    '============================================================
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Me.Icon = LoadResPicture(101, 1)
    DoGradient Me
End Sub

Public Function GetViewData(SearchText As String)
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim strSearch
    Set db = OpenDatabase(DBPath)
    Set rsCode = db.OpenRecordset("tblProjectCode", dbOpenDynaset)
    strSearch = "[Codename] Like '" & SearchText & "'"
    With rsCode
        .FindFirst strSearch
        txtAuthor.Text = rsCode.Fields("Author")
        txtPhone.Text = rsCode.Fields("Phone")
        txtCodeName.Text = rsCode.Fields("CodeName")
        txtPurpose.Text = rsCode.Fields("Purpose")
        txtRelDate.Text = rsCode.Fields("DateReleased")
        rtfCode.Text = rsCode.Fields("SourceCode")
        Me.Caption = "Source Code for: " & rsCode.Fields("CodeName")
    End With
End Function

Private Sub cmdCopy_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Clipboard.Clear
    Clipboard.SetText (rtfCode.Text)
    MsgBox "Code copied to clipboard.  You may now paste it into your project.", vbOKOnly + vbInformation, "Code Copied!"
End Sub

