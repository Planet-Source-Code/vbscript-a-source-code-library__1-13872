VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit/Update Source Code"
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
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Code"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Update/Edit current code."
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      ToolTipText     =   "Clear changes."
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      ToolTipText     =   "Close form."
      Top             =   120
      Width           =   975
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
   Begin VB.ComboBox cmbLanguage 
      Height          =   315
      Left            =   6720
      TabIndex        =   6
      Top             =   960
      Width           =   1695
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
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   3000
      MaxLength       =   12
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox cmbAuthor 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtOrigDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtUpdateDate 
      Alignment       =   2  'Center
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Width           =   1095
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
      ScrollBars      =   2
      TextRTF         =   $"frmEdit.frx":0000
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Author"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Phone/Extension"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblCodeName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code Name/Title"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lblOrigDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Create Date"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblUpdateDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Update"
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   720
      Width           =   1095
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
   Begin VB.Label lblLanguage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Language"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblPurpose 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purpose/Function of the Code"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Source Code - Cut and paste into your document"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Label lblKeywords 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Searchable Keywords"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '============================================================
    ' This allows a record to be edited.  This section sets up
    ' the global variables used throughout the form.
    '============================================================
    Dim db As Database
    Dim rsCode As Recordset
    Dim rsAuth As Recordset
    Dim rsLang As Recordset

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    ' This runs everytime the form is loaded.  It sets the form
    ' icon, sets the form's background to a gradient, opens the
    ' database and tables, populates the Author dropdown box and
    ' Language dropdown box, and sets the UpdateDate field to
    ' the current date.
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
    txtUpdateDate.Text = Date
End Sub

Private Sub cmdReset_Click()
On Error Resume Next
    '============================================================
    ' This resets the fields to what is currently in the
    ' database.
    '============================================================
    GetEditData (txtCodeName.Text)
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
    ' This syncronizes the Author name with the Author phone
    ' number.  When an Author's name is selected, the Phone field
    ' updaes with the matching Phone number.
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

Private Sub cmdUpdate_Click()
On Error Resume Next
    '============================================================
    ' Once all changes are entered, this varifies user input and
    ' updates the record.
    '============================================================
    With rsCode
        .Edit
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
    GetEditData (txtCodeName.Text)
End Sub

Public Function GetEditData(SearchText As String)
On Error Resume Next
    '============================================================
    ' This loads the default data based on the search string sent
    ' from the selected listbox entry.  It populates the fields
    ' based on what is currently in the database.
    '============================================================
    Dim strSearch
    Set db = OpenDatabase(DBPath)
    Set rsCode = db.OpenRecordset("tblSourceCode", dbOpenDynaset)
    strSearch = "[CodeName] = '" & SearchText & "'"
    rsCode.MoveFirst
    With rsCode
        .FindFirst strSearch
        cmbAuthor.Text = rsCode.Fields("Author")
        txtPhone.Text = rsCode.Fields("Phone")
        txtCodeName.Text = rsCode.Fields("CodeName")
        txtPurpose.Text = rsCode.Fields("Purpose")
        txtOrigDate.Text = rsCode.Fields("OrigDate")
        txtUpdateDate.Text = Date
        txtVersion.Text = rsCode.Fields("Version")
        cmbLanguage.Text = rsCode.Fields("Language")
        txtKeywords.Text = rsCode.Fields("Keywords")
        rtfCode.Text = rsCode.Fields("Code")
        Me.Caption = "Source Code for: " & rsCode.Fields("CodeName")
    End With
End Function
