VERSION 5.00
Begin VB.Form frmEditLang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Languages"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Language"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Edit highlighted language."
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Um, close the form."
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstLang 
      Height          =   1425
      ItemData        =   "frmEditLang.frx":0000
      Left            =   120
      List            =   "frmEditLang.frx":0002
      TabIndex        =   3
      ToolTipText     =   "Displays the language name."
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox txtLang 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Edit box for language name."
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Language"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Add new language to the database."
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '============================================================
    ' This window allows the user to update and add records to
    ' the Language Table.  This section sets up the global
    ' variables used throughout the form.
    '============================================================
    Dim db As Database
    Dim rsLang As Recordset

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    ' This is run everytime the window is opened.  It sets the
    ' window icon, runs the Gradient function and calls the sub
    ' to populate the listbox.
    '============================================================
    Me.Icon = LoadResPicture(111, 1)
    DoGradient Me
    DoRefresh
End Sub

Private Sub DoRefresh()
On Error Resume Next
    '============================================================
    ' This opens the database and table.  It populates the
    ' listbox from the Languages table.
    '============================================================
    Set db = OpenDatabase(DBPath)
    Set rsLang = db.OpenRecordset("tblLanguage", dbOpenDynaset)
    lstLang.Clear
    txtLang.Text = ""
    rsLang.MoveFirst
    Do While Not rsLang.EOF
        lstLang.AddItem (rsLang.Fields("Language"))
        rsLang.MoveNext
    Loop
End Sub
    
Private Sub cmdClose_Click()
On Error Resume Next
    '============================================================
    ' This closes the window.
    '============================================================
    Unload Me
End Sub

Private Sub lstLang_Click()
On Error Resume Next
    '============================================================
    ' This selects the record for editing.
    '============================================================
    Dim intLoopIndex As Integer
    intLoopIndex = 0
    For intLoopIndex = 0 To lstLang.ListCount - 1
        If lstLang.Selected(intLoopIndex) Then
            txtLang.Text = lstLang.Text
        End If
    Next intLoopIndex
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
    '============================================================
    ' This adds a new record in the Language table.
    '============================================================
    With rsLang
        .AddNew
        !Language = txtLang.Text
        .Update
    End With
    DoRefresh
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
    '============================================================
    ' This edits the selected record in the Language table.
    '============================================================
    rsLang.MoveFirst
    Do Until rsLang.EOF
        If txtLang.Text = rsLang.Fields("Language") Then
            With rsLang
                .Edit
                !Language = txtLang.Text
                .Update
            End With
            DoRefresh
        Else
            rsLang.MoveNext
        End If
    Loop
    DoRefresh
End Sub


