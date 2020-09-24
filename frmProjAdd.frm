VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProjAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Project Code"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   8535
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Code"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "Add new code to the database"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      ToolTipText     =   "Clear current information"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      ToolTipText     =   "Close form"
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbAuthor 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtPhone 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtCodeName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtRelDate 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtPurpose 
      Height          =   495
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   840
      Width           =   3735
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4471
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmProjAdd.frx":0000
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Author"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2775
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
   Begin VB.Label lblCodeName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Project Title"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblOrigDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Release Date"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblPurpose 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purpose/Function of the Project"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Source Code - Cut and paste into your document"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   8295
   End
End
Attribute VB_Name = "frmProjAdd"
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
    Dim rsAuth As Recordset

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Me.Icon = LoadResPicture(103, 1)
    DoGradient Me
    Set db = OpenDatabase(DBPath)
    Set rsCode = db.OpenRecordset("tblProjectCode", dbOpenDynaset)
    Set rsAuth = db.OpenRecordset("tblAuthor", dbOpenDynaset)
    rsAuth.MoveFirst
    Do While Not rsAuth.EOF
        cmbAuthor.AddItem (rsAuth.Fields("Author"))
        rsAuth.MoveNext
    Loop
    txtRelDate.Text = Date
End Sub

Private Sub cmdAdd_Click()
    '============================================================
    '
    '============================================================
    With rsCode
        .AddNew
        If cmbAuthor.Text = "" Then
            !Author = "No one Listed"
        Else
            !Author = cmbAuthor.Text
        End If
        If txtPhone.Text = "" Then
            !Phone = "No Phone"
        Else
            !Phone = txtPhone.Text
        End If
        If txtCodeName.Text = "" Then
            !CodeName = InputBox("Project Name is REQUIRED.  Please enter a name for this Project.", "Input Required!!!")
        Else
            !CodeName = txtCodeName.Text
        End If
        If txtPurpose.Text = "" Then
            !Purpose = "This code has no purpose"
        Else
            !Purpose = txtPurpose.Text
        End If
        If txtRelDate = "" Then
            !DateReleased = Date
        Else
            !DateReleased = txtRelDate.Text
        End If
        If rtfCode.Text = "" Then
            !SourceCode = "No Code Listed"
        Else
            !SourceCode = rtfCode.Text
        End If
        !Project = !CodeName
        .Update
    End With
    cmdReset_Click
End Sub

Private Sub cmdReset_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    cmbAuthor.Text = ""
    txtPhone.Text = ""
    txtCodeName.Text = ""
    txtRelDate.Text = Date
    txtPurpose.Text = ""
    rtfCode.Text = ""
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Unload Me
End Sub

Private Sub cmbAuthor_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim strSearch As String
    If cmbAuthor.Text = "" Then
        Exit Sub
    Else
        strSearch = "[Author] Like '" & cmbAuthor.Text & "'"
        With rsAuth
            .FindFirst strSearch
            txtPhone.Text = rsAuth.Fields("Phone")
        End With
    End If
End Sub

