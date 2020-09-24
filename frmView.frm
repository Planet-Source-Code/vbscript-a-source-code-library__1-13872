VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Source Code"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   8535
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Send Code to Clipboard"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      ToolTipText     =   "Copies full text from the Source Code window into the clipboard to be pasted in your application."
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtKeywords 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txtLanguage 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtAuthor 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtPurpose 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      ToolTipText     =   "Close current form"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtOrigDate 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtCodeName 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtUpdateDate 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtPhone 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3201
      _Version        =   393217
      BackColor       =   12648447
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmView.frx":0000
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Source Code - Cut and paste into your document"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Label lblKeywords 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Searchable Keywords"
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblPurpose 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purpose/Function of the Code"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblLanguage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Language"
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Version"
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblUpdateDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Update"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblOrigDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Create Date"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblCodeName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code Name/Title"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Phone/Extension"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Author"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmView"
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

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Me.Icon = LoadResPicture(101, 1)
    DoGradient Me
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Unload Me
End Sub

Public Function GetViewData(SearchText As String)
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim strSearch
    Set db = OpenDatabase(DBPath)
    Set rsCode = db.OpenRecordset("tblSourceCode", dbOpenDynaset)
    strSearch = "[Codename] Like '" & SearchText & "'"
    With rsCode
        .FindFirst strSearch
        txtAuthor.Text = rsCode.Fields("Author")
        txtPhone.Text = rsCode.Fields("Phone")
        txtCodeName.Text = rsCode.Fields("CodeName")
        txtPurpose.Text = rsCode.Fields("Purpose")
        txtOrigDate.Text = rsCode.Fields("OrigDate")
        txtUpdateDate.Text = rsCode.Fields("UpdateDate")
        txtVersion.Text = rsCode.Fields("Version")
        txtLanguage.Text = rsCode.Fields("Language")
        txtKeywords.Text = rsCode.Fields("Keywords")
        rtfCode.Text = rsCode.Fields("Code")
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
