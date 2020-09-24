VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Source Code"
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
   Begin VB.TextBox txtKeywords 
      Height          =   495
      Left            =   4320
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1560
      Width           =   4095
   End
   Begin VB.ComboBox cmbAuthor 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   3000
      MaxLength       =   12
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtUpdateDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtCodeName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtOrigDate 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cmbLanguage 
      Height          =   315
      Left            =   6720
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5880
      MaxLength       =   5
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      ToolTipText     =   "Close form"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      ToolTipText     =   "Clear current information"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Code"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Add new code to the database"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPurpose 
      Height          =   495
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1560
      Width           =   4095
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAdd.frx":0000
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Author"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Phone/Extension"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblCodeName 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code Name/Title"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lblOrigDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Create Date"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblUpdateDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Update"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Version"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLanguage 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Language"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblPurpose 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purpose/Function of the Code"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Source Code - Cut and paste into your document"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Label lblKeywords 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Searchable Keywords"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '============================================================
    ' This form adds a record into the SourceCode table.
    ' This section sets the global variables used throughout this
    ' form.
    '============================================================
    Dim db As Database
    Dim rsCode As Recordset
    Dim rsAuth As Recordset
    Dim rsLang As Recordset

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    ' This section runs everytime the form is loaded.  It first
    ' sets up the window's icon.  Next, calls the function to set
    ' up the form's background to a gradient.  Next, it opens the
    ' database and loads the tables.  Then, it populates the
    ' Author's dropdown list and then the Language dropdown list.
    ' It also sets the default values for the two date fields and
    ' the version field.
    '============================================================
    Me.Icon = LoadResPicture(103, 1)
    DoGradient Me
    Set db = OpenDatabase(DBPath)
    Set rsCode = db.OpenRecordset("tblSourceCode", dbOpenDynaset)
    Set rsAuth = db.OpenRecordset("tblAuthor", dbOpenDynaset)
    Set rsLang = db.OpenRecordset("tblLanguage", dbOpenDynaset)
    rsAuth.MoveFirst
    Do While Not rsAuth.EOF
        cmbAuthor.AddItem (rsAuth.Fields("Author"))
        rsAuth.MoveNext
    Loop
    rsLang.MoveFirst
    Do While Not rsLang.EOF
        cmbLanguage.AddItem (rsLang.Fields("Language"))
        rsLang.MoveNext
    Loop
    txtOrigDate.Text = Date
    txtUpdateDate.Text = Date
    txtVersion.Text = "1.00"
End Sub

Private Sub cmdReset_Click()
On Error Resume Next
    '============================================================
    ' This clears the form completly, reseting the defaults.
    '============================================================
    cmbAuthor.Text = ""
    txtPhone.Text = ""
    txtCodeName.Text = ""
    txtOrigDate.Text = Date
    txtUpdateDate.Text = Date
    txtVersion.Text = "1.00"
    cmbLanguage.Text = ""
    txtPurpose.Text = ""
    txtKeywords.Text = ""
    rtfCode.Text = ""
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
    '============================================================
    ' This closes the window.
    '============================================================
    Unload Me
End Sub

Private Sub cmbAuthor_Click()
On Error Resume Next
    '============================================================
    ' This function syncronizes the Author Name field and the
    ' Author Phone/Extension field, making sure they match up
    ' when an Author is selected.
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

Private Sub cmdAdd_Click()
On Error Resume Next
    '============================================================
    ' When the Add Code button is clicked, this sub is called.
    ' It checks the values entered for NULL and fills in the ones
    ' left blank.  The only field that requires user input is the
    ' Code Name.  If that field is left blank, it pops up an
    ' inputbox requiring user input.  After the record is entered
    ' the form is cleared and readied for another entry.
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
            !CodeName = InputBox("Code Name is REQUIRED.  Please a name for this code snippet.", "Input Required!!!")
        Else
            !CodeName = txtCodeName.Text
        End If
        If txtPurpose.Text = "" Then
            !Purpose = "This code has no purpose"
        Else
            !Purpose = txtPurpose.Text
        End If
        If txtOrigDate = "" Then
            !OrigDate = Date
        Else
            !OrigDate = txtOrigDate.Text
        End If
        If txtUpdateDate.Text = "" Then
            !UpdateDate = Date
        Else
            !UpdateDate = txtUpdateDate.Text
        End If
        If txtVersion = "" Then
            !Version = "1.00"
        Else
            !Version = txtVersion.Text
        End If
        If cmbLanguage.Text = "" Then
            !Language = "No Language Listed"
        Else
            !Language = cmbLanguage.Text
        End If
        If txtKeywords.Text = "" Then
            !Keywords = "None"
        Else
            !Keywords = txtKeywords.Text
        End If
        If rtfCode.Text = "" Then
            !Code = "No Code Listed"
        Else
            !Code = rtfCode.Text
        End If
        .Update
    End With
    cmdReset_Click
End Sub
