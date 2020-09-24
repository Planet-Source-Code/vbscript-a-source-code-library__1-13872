VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for Source Code"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   5535
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      ToolTipText     =   "Reset/Clear the results and criteria."
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      ToolTipText     =   "Edit highlighted code."
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   "View highlighted code."
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Do the search"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CheckBox chkKeyword 
      Caption         =   "Check3"
      Height          =   190
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Enable/Disable Search by Keywords"
      Top             =   840
      Width           =   215
   End
   Begin VB.CheckBox chkLanguage 
      Height          =   190
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Enable/Disable Search by Language"
      Top             =   480
      Width           =   215
   End
   Begin VB.CheckBox chkAuthor 
      Height          =   190
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enable/Disable Search by Author"
      Top             =   120
      Width           =   215
   End
   Begin VB.TextBox txtKeyword 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      ToolTipText     =   "Type in up to 5 COMMA seperated keywords."
      Top             =   840
      Width           =   3255
   End
   Begin VB.ComboBox cmbLanguage 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "This list shows the current Languages to search by."
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox cmbAuthor 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "This list shows the current authors to search by"
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox lstResults 
      Height          =   1815
      ItemData        =   "frmSearch.frx":0000
      Left            =   120
      List            =   "frmSearch.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Here are your results."
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label lblKeyword 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Keywords"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      ToolTipText     =   "Enable/Disable Search by Keywords"
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblLanguage 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Language"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "Enable/Disable Search by Language"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Author"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      ToolTipText     =   "Enable/Disable Search by Author"
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '============================================================
    '
    '============================================================
    Dim strKeywords As String
    Dim db As Database
    Dim rsCode As Recordset
    Dim rsAuth As Recordset
    Dim rsLang As Recordset
    Dim ListType As String

Private Sub Form_Load()
On Error Resume Next
    '============================================================
    '
    '============================================================
    ListType = GetSetting("SourceCode", "Settings", "DataSet")
    Me.Icon = LoadResPicture(105, 1)
    Set db = OpenDatabase(DBPath)
    Select Case ListType
        Case "Source"
            Set rsCode = db.OpenRecordset("tblSourceCode", dbOpenDynaset)
            Set rsLang = db.OpenRecordset("tblLanguage", dbOpenDynaset)
            rsLang.MoveFirst
            Do While Not rsLang.EOF
                cmbLanguage.AddItem (rsLang.Fields("Language"))
                rsLang.MoveNext
            Loop
        Case "Project"
            Set rsCode = db.OpenRecordset("tblProjectCode", dbOpenDynaset)
    End Select
    Set rsAuth = db.OpenRecordset("tblAuthor", dbOpenDynaset)
    DoGradient Me
    rsAuth.MoveFirst
    Do While Not rsAuth.EOF
        cmbAuthor.AddItem (rsAuth.Fields("Author"))
        rsAuth.MoveNext
    Loop
    cmdReset_Click
End Sub

Private Sub cmdReset_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    ListType = GetSetting("SourceCode", "Settings", "DataSet")
    Select Case ListType
        Case "Source"
            lstResults.Clear
            cmbAuthor.Enabled = False
            cmbLanguage.Enabled = False
            txtKeyword.Enabled = False
            chkAuthor.Value = 0
            chkLanguage.Value = 0
            chkKeyword.Value = 0
            cmbAuthor.Text = ""
            cmbLanguage.Text = ""
            txtKeyword.Text = ""
            txtKeyword.BackColor = &H80000004
        Case "Project"
            lstResults.Clear
            cmbAuthor.Enabled = False
            cmbLanguage.Enabled = False
            txtKeyword.Enabled = False
            chkAuthor.Value = 0
            chkLanguage.Value = 0
            chkLanguage.Enabled = False
            chkKeyword.Value = 0
            chkKeyword.Enabled = False
            cmbAuthor.Text = ""
            cmbLanguage.Text = ""
            txtKeyword.Text = ""
            txtKeyword.BackColor = &H80000004
            cmdEdit.Enabled = False
    End Select
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim intLoopIndex As Integer
    intLoopIndex = 0
    For intLoopIndex = 0 To lstResults.ListCount - 1
        If lstResults.Selected(intLoopIndex) Then
            frmEdit.GetEditData lstResults.Text
        End If
    Next intLoopIndex
    frmEdit.Show
End Sub

Private Sub cmdView_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim intLoopIndex As Integer
    intLoopIndex = 0
    For intLoopIndex = 0 To lstResults.ListCount - 1
        If lstResults.Selected(intLoopIndex) Then
            Select Case ListType
                Case "Source"
                    frmView.GetViewData lstResults.Text
                Case "Project"
                    frmProjView.GetViewData lstResults.Text
            End Select
        End If
    Next intLoopIndex
    Select Case ListType
        Case "Source"
            frmView.Show
        Case "Project"
            frmProjView.Show
    End Select
End Sub

Private Sub chkAuthor_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    If chkAuthor.Value = 1 Then
        cmbAuthor.Enabled = True
    ElseIf chkAuthor.Value = 0 Then
        cmbAuthor.Enabled = False
    End If
End Sub

Private Sub chkLanguage_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    If chkLanguage.Value = 1 Then
        cmbLanguage.Enabled = True
    ElseIf chkLanguage.Value = 0 Then
        cmbLanguage.Enabled = False
    End If
End Sub

Private Sub chkKeyword_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    If chkKeyword.Value = 1 Then
        txtKeyword.Enabled = True
        txtKeyword.BackColor = &H80000005
    ElseIf chkKeyword.Value = 0 Then
        txtKeyword.Enabled = False
        txtKeyword.BackColor = &H80000004
    End If
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim strSearch As String
    Dim strAuthor As String
    Dim strLanguage As String
    
    strSearch = ""
    strAuthor = ""
    strKeywords = ""
    
    lstResults.Clear
    
    strAuthor = "[Author] Like '" & cmbAuthor.Text & "'"
    strLanguage = "[Language] Like '" & cmbLanguage.Text & "'"
    
    If txtKeyword.Text > "" Then
        GetKeywords
    End If
    
    If chkAuthor.Value = 1 And chkLanguage.Value = 0 And chkKeyword.Value = 0 Then
        strSearch = strAuthor
    ElseIf chkAuthor.Value = 0 And chkLanguage.Value = 1 And chkKeyword.Value = 0 Then
        strSearch = strLanguage
    ElseIf chkAuthor.Value = 0 And chkLanguage.Value = 0 And chkKeyword.Value = 1 Then
        strSearch = strKeywords
    ElseIf chkAuthor.Value = 1 And chkLanguage.Value = 1 And chkKeyword.Value = 0 Then
        strSearch = strAuthor & " AND " & strLanguage
    ElseIf chkAuthor.Value = 0 And chkLanguage.Value = 1 And chkKeyword.Value = 1 Then
        strSearch = strLanguage & " AND " & strKeywords
    ElseIf chkAuthor.Value = 1 And chkLanguage.Value = 0 And chkKeyword.Value = 1 Then
        strSearch = strAuthor & " AND " & strKeywords
    ElseIf chkAuthor.Value = 1 And chkLanguage.Value = 1 And chkKeyword.Value = 1 Then
        strSearch = strAuthor & " AND " & strLanguage & " AND " & strKeywords
    End If
    
    With rsCode
        .FindFirst strSearch
        If .NoMatch Then
            MsgBox "Your search is complete.", vbInformation + vbOKOnly, "Search Complete"
        Else
            lstResults.AddItem (rsCode.Fields("CodeName"))
            Again strSearch
        End If
    End With
End Sub

Private Sub cmdExit_Click()
    '============================================================
    '
    '============================================================
    Unload Me
End Sub

Private Sub Again(strSearch As String)
On Error Resume Next
    '============================================================
    '
    '============================================================
    With rsCode
        .FindNext strSearch
        If .NoMatch Then
            MsgBox "Your search is complete.", vbInformation + vbOKOnly, "Search Complete"
        Else
            lstResults.AddItem (rsCode.Fields("CodeName"))
            Again strSearch
        End If
    End With
End Sub

Private Sub lstResults_DblClick()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim intLoopIndex As Integer
    intLoopIndex = 0
    For intLoopIndex = 0 To lstResults.ListCount - 1
        If lstResults.Selected(intLoopIndex) Then
            Select Case ListType
                Case "Source"
                    frmView.GetViewData lstResults.Text
                Case "Project"
                    frmProjView.GetViewData lstResults.Text
            End Select
        End If
    Next intLoopIndex
    Select Case ListType
        Case "Source"
            frmView.Show
        Case "Project"
            frmProjView.Show
    End Select
End Sub

Public Sub GetKeywords()
On Error Resume Next
    '============================================================
    '
    '============================================================
    Dim Source() As String
    Dim x As Integer
    Dim Textnum As Integer
    Source() = Split(txtKeyword.Text, ",", 5)
    If UBound(Source()) = 0 Then
        strKeywords = "[Keywords] Like '*" & Source(0) & "*'"
    Else
        For x = 0 To UBound(Source())
            If Source(x) = "" Then
                strKeywords = "[Keywords] Like '*" & Source(x) & "*'"
            Else
                strKeywords = "[Keywords] Like '*" & Source(x) & "*' OR " + strKeywords
             End If
        Next x
        Textnum = Len(strKeywords)
        strKeywords = Left(strKeywords, Textnum - 4)
    End If
End Sub
