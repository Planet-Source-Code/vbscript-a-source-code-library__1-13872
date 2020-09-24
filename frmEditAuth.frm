VERSION 5.00
Begin VB.Form frmEditAuth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Authors"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Author"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Add new author to database"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "This is the edit box for the author's phone number."
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtAuth 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "This is the edit box for the author's name"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ListBox lstPhone 
      Height          =   1425
      ItemData        =   "frmEditAuth.frx":0000
      Left            =   2040
      List            =   "frmEditAuth.frx":0002
      TabIndex        =   4
      ToolTipText     =   "This pane displays the authors' phone number"
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox lstAuth 
      Height          =   1425
      ItemData        =   "frmEditAuth.frx":0004
      Left            =   120
      List            =   "frmEditAuth.frx":0006
      TabIndex        =   3
      ToolTipText     =   "This pane displays the author's name."
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Um, close this form."
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Author"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Edit highlighted author"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '============================================================
    ' This is the form that allows the Author Table to be edited.
    ' This section sets up the global variables used throughout
    ' the form.
    '============================================================
    Dim db As Database
    Dim rsAuth As Recordset

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    ' This section runs everytime the form is loaded.  IIt sets
    ' the window icon, runs the Gradient function, and populates
    ' the listboxes.
    '============================================================
    Me.Icon = LoadResPicture(111, 1)
    DoGradient Me
    DoRefresh
End Sub

Private Sub DoRefresh()
On Error Resume Next
    '============================================================
    ' This clears the edit boxes, and populates the two listboxes
    ' with the data from the Author Table.
    '============================================================
    Set db = OpenDatabase(DBPath)
    Set rsAuth = db.OpenRecordset("tblAuthor", dbOpenDynaset)
    lstAuth.Clear
    lstPhone.Clear
    txtAuth.Text = ""
    txtPhone.Text = ""
    rsAuth.MoveFirst
    Do While Not rsAuth.EOF
        lstAuth.AddItem (rsAuth.Fields("Author"))
        lstPhone.AddItem (rsAuth.Fields("Phone"))
        rsAuth.MoveNext
    Loop
End Sub
    
Private Sub cmdClose_Click()
On Error Resume Next
    '============================================================
    ' This closes the window.
    '============================================================
    Unload Me
End Sub

Private Sub lstAuth_Click()
On Error Resume Next
    '============================================================
    ' This syncronizes the Author Name listbox and Author Phone
    ' listbox.  Also, when an author is selected, it fills in the
    ' fields at the bottom of the form.
    '============================================================
    lstPhone.ListIndex = lstAuth.ListIndex
    Dim intLoopIndex As Integer
    intLoopIndex = 0
    For intLoopIndex = 0 To lstAuth.ListCount - 1
        If lstAuth.Selected(intLoopIndex) Then
            txtAuth.Text = lstAuth.Text
            txtPhone.Text = lstPhone.Text
        End If
    Next intLoopIndex
End Sub

Private Sub lstPhone_Click()
On Error Resume Next
    '============================================================
    ' This does the same as the above sub.
    '============================================================
    lstAuth.ListIndex = lstPhone.ListIndex
    Dim intLoopIndex As Integer
    intLoopIndex = 0
    For intLoopIndex = 0 To lstAuth.ListCount - 1
        If lstAuth.Selected(intLoopIndex) Then
            txtAuth.Text = lstAuth.Text
            txtPhone.Text = lstPhone.Text
        End If
    Next intLoopIndex
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
    '============================================================
    ' This adds a new author record to the database
    '============================================================
    With rsAuth
        .AddNew
        !Author = txtAuth.Text
        !Phone = txtPhone.Text
        .Update
    End With
    DoRefresh
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
    '============================================================
    ' This edits an existing author record in the database
    '============================================================
    rsAuth.MoveFirst
    Do Until rsAuth.EOF
        If txtAuth.Text = rsAuth.Fields("Author") Then
            With rsAuth
                .Edit
                !Author = txtAuth.Text
                !Phone = txtPhone.Text
                .Update
            End With
            DoRefresh
        Else
            rsAuth.MoveNext
        End If
    Loop
    DoRefresh
End Sub
