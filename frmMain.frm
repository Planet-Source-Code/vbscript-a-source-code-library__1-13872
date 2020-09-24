VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Humana Source Code Library"
   ClientHeight    =   7395
   ClientLeft      =   255
   ClientTop       =   825
   ClientWidth     =   12135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":030A
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picList 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6540
      Left            =   0
      ScaleHeight     =   6540
      ScaleWidth      =   2295
      TabIndex        =   2
      Top             =   600
      Width           =   2295
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh from Database"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2055
      End
      Begin VB.ListBox lstCode 
         BackColor       =   &H00C0FFFF&
         Height          =   5325
         ItemData        =   "frmMain.frx":927C
         Left            =   120
         List            =   "frmMain.frx":927E
         Sorted          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "This pane displays the code title."
         Top             =   1080
         Width           =   2040
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "View highlighted code."
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Edit highlighted code."
         Top             =   600
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList LargeImages 
      Left            =   2880
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9280
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":959A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9E74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A18E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A4A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B09C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B976
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C250
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CB2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D404
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D71E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DA38
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD52
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E06C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E386
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E6A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E9BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbTools 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "LargeImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Opens a new database."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Settings"
            Object.ToolTipText     =   "Opens the Settings Dialog"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit the Program"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "Add new Code Snippet"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "List"
            Object.ToolTipText     =   "List current Code Snippets"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Search for specific example"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Project"
            Object.ToolTipText     =   "Show Project Code"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Source"
            Object.ToolTipText     =   "Show Code Snippets"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Author"
            Object.ToolTipText     =   "Update/Change Author Information"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Lang"
            Object.ToolTipText     =   "Update/Change Language Information"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cascade"
            Object.ToolTipText     =   "Arrange Icons"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileH"
            Object.ToolTipText     =   "Tile Windows Horizontally"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileV"
            Object.ToolTipText     =   "Tile Windows Vertically"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Arrange"
            Object.ToolTipText     =   "Arrange Icons"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Display Help File"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "Display About Dialog"
            ImageIndex      =   18
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7140
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15769
            Object.ToolTipText     =   "Shows current number of Code Examples in the database."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "5:34 PM"
            Object.ToolTipText     =   "Displays the current time."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/30/2000"
            Object.ToolTipText     =   "Displays the current date."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Database"
      Filter          =   "*.mdb,*.dat"
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   3480
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ECD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EFEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F8C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FBE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FEFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":107D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":113CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1257E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13172
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1348C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":137A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":140F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1440E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSettings 
         Caption         =   "&Settings"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Database"
      Begin VB.Menu mnuDataSourceList 
         Caption         =   "&Code Snippets"
      End
      Begin VB.Menu mnuDataProjList 
         Caption         =   "&Project Code"
      End
      Begin VB.Menu mnuDataSpace01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataSearch 
         Caption         =   "&Search Code"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuDataAdd 
         Caption         =   "&Add New Code"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDataSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMain 
         Caption         =   "&Maintainance"
         Begin VB.Menu mnuMainAuth 
            Caption         =   "&Add/Change Authors"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuMainLang 
            Caption         =   "Add/&Change Language"
            Shortcut        =   ^L
         End
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHoriz 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVert 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "Window List"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '============================================================
    ' This is the main form of the program.
    ' Here the global variables that are used throughout this
    ' form are set up.
    '============================================================
    Dim db As Database
    Dim rsCode As Recordset
    Dim ListType As String
    
Public Sub MDIForm_Load()
On Error Resume Next
    '============================================================
    ' This function runs whenever the form is loaded or reloaded.
    ' The first thing that happend is it loads the correct icon
    ' then it sets the default view for the listbox, finally it
    ' calls the sub for refreshing the listbox.
    '============================================================
    Dim lngVersion As Long, TransState As String, TransDeg As String
    lngVersion = GetVersion()
    frmMain.Icon = LoadResPicture(120, 1)
    TransState = GetSetting("SourceCode", "Settings", "TransState", "0")
    TransDeg = GetSetting("SourceCode", "Settings", "TransDeg", "128")
    Select Case TransState
        Case 0
            If lngVersion = 143851525 Then
                SetLayered frmMain.hWnd, False, 255
            End If
        Case 1
            If lngVersion = 143851525 Then
                SetLayered frmMain.hWnd, True, Int(TransDeg)
            End If
    End Select
    SaveSetting "SourceCode", "Settings", "DataSet", "Source"
    cmdRefresh_Click
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    '============================================================
    ' This makes sure the window cannot be maximized.  It checks
    ' when the window is resized, the windowstate.  Based on that
    ' it either does nothing, resets the state, or changes the
    ' size back to the default setting.  No, not the best method,
    ' but it works, if a bit crude.
    '============================================================
    If Me.WindowState = 1 Then
        'Me.WindowState = 1
    ElseIf Me.WindowState = 2 Then
        Me.WindowState = 0
    Else
        Me.Height = 8085
        Me.Width = 12255
    End If
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
    '============================================================
    ' This does a little more than refresh the list.  First, it
    ' retrieves two settings from the registry, Toolsize - for
    ' the toolbar button size, and ListType - for the correct
    ' table to read from i nthe source database.
    '
    ' Based on that information, the Select Case statement sets
    ' up the correct image list as the default for the toolbar.
    ' After it sets the imagelist property for the toolbar, it
    ' calls the GetButtons sub to set the images property for
    ' each button.
    '
    ' After it finishes up setting the images, it comes back and
    ' clears the list.  Then the sub GetData is called and sent
    ' the ListType value.
    '============================================================
    Dim Count, ToolSize, ListAlign
    ToolSize = GetSetting("SourceCode", "Settings", "Icons", "0")
    ListType = GetSetting("SourceCode", "Settings", "DataSet", "Source")
    ListAlign = GetSetting("SourceCode", "Settings", "AlignList")
    Select Case ListAlign
        Case 3
            picList.Align = 3
        Case 4
            picList.Align = 4
    End Select
    Select Case ToolSize
        Case 0
            tbTools.ImageList = LargeImages
            GetButtons
        Case 1
            tbTools.ImageList = SmallImages
            GetButtons
    End Select
    lstCode.Clear
    GetData ListType
End Sub

Private Sub cmdView_Click()
On Error Resume Next
    '============================================================
    ' Again, the ListType variable is set from the Registry.  The
    ' first thing that happens is the selected listbox item is
    ' verified, and the selected item's text is sent to the
    ' appropriate GetViewData for the ListType.
    '
    ' Then it shows the appropriate form based on the ListType
    ' setting.
    '============================================================
    Dim intLoopIndex As Integer
    ListType = GetSetting("SourceCode", "Settings", "DataSet")
    intLoopIndex = 0
    For intLoopIndex = 0 To lstCode.ListCount - 1
        If lstCode.Selected(intLoopIndex) Then
            Select Case ListType
                Case "Source"
                    frmView.GetViewData lstCode.Text
                Case "Project"
                    frmProjView.GetViewData lstCode.Text
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

Private Sub lstCode_DblClick()
On Error Resume Next
    '============================================================
    ' This sub does nothing except call the View sub whenever a
    ' list item is double clicked.
    '============================================================
    cmdView_Click
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
    '============================================================
    ' This sub is only used when the ListType is "Source".  The
    ' Edit button itself is disabled in the "Project" View.  The
    ' sub takes the selected item and sends it's text to the Edit
    ' form's GetEditData.  After the Edit form's fields are
    ' populated, it shows the Edit Form.
    '============================================================
    Dim intLoopIndex As Integer
    intLoopIndex = 0
    For intLoopIndex = 0 To lstCode.ListCount - 1
        If lstCode.Selected(intLoopIndex) Then
            frmEdit.GetEditData lstCode.Text
        End If
    Next intLoopIndex
    frmEdit.Show
End Sub

Private Sub mnuDataAdd_Click()
On Error Resume Next
    '============================================================
    ' Based on the ListType setting, this opens the appropriate
    ' form to add code to the database.
    '============================================================
    ListType = GetSetting("SourceCode", "Settings", "DataSet")
    Select Case ListType
        Case "Source"
            frmAdd.Show
        Case "Project"
            frmProjAdd.Show
    End Select
End Sub

Private Sub mnuDataProjList_Click()
    '============================================================
    ' This sets the DataSet setting in the registry to "Project."
    ' This setting is used to set the ListType throughout the
    ' pogram.  This sub also disables the Edit button and calls
    ' the Refresh sub to refresh the listbox with the appropriate
    ' data.
    '============================================================
    SaveSetting "SourceCode", "Settings", "DataSet", "Project"
    cmdEdit.Enabled = False
    cmdRefresh_Click
End Sub

Private Sub mnuDataSearch_Click()
On Error Resume Next
    '============================================================
    ' This command shows the Serch Form Window.
    '============================================================
    frmSearch.Show
End Sub

Private Sub mnuDataSourceList_Click()
On Error Resume Next
    '============================================================
    ' This sets the DataSet setting in the registry to "Source."
    ' This setting is used to set the ListType throughout the
    ' pogram.  This sub also enables the Edit button and calls
    ' the Refresh sub to refresh the listbox with the appropriate
    ' data.
    '============================================================
    SaveSetting "SourceCode", "Settings", "DataSet", "Source"
    cmdEdit.Enabled = True
    cmdRefresh_Click
End Sub

Private Sub mnuFileExit_Click()
On Error Resume Next
    '============================================================
    ' This command exits the program.
    '============================================================
    Unload Me
    End
End Sub

Private Sub mnuFileOpen_Click()
On Error Resume Next
    '============================================================
    ' This sub is used to open the common dialog box for opening
    ' the database, or switching databases.
    '
    ' The first thing it does is get the current path from the
    ' registry setting.  It then sets up the default properties
    ' for the common dialog box.  It sets the filter to show only
    ' .mdb and .dat files, and then it saves the new selection in
    ' in the registry.  Finally, it reloads the form.
    '============================================================
    Dim sFile As String
    sFile = GetSetting("SourceCode", "Settings", "Path")
    With dlgOpen
        .DialogTitle = "Open Database"
        .CancelError = False
        .Filter = "Database Files (*.dat,*.mdb)|*.dat;*.mdb|"
        .Filter = .Filter + "Access Databases (*.mdb)|*.mdb|"
        .Filter = .Filter + "Dat Files (*.dat)|*.dat|"
        .Filter = .Filter + "All Files (*.*)|*.*"
        .InitDir = sFile
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    SaveSetting "SourceCode", "Settings", "Path", sFile
    MDIForm_Load
End Sub

Private Sub mnuFileSettings_Click()
On Error Resume Next
    '============================================================
    ' This opens the settings window and fills out the default
    ' information from the settings in the registry.  It then
    ' opens the settins window.
    '============================================================
    Dim IconSize, Trans, ListDock
    IconSize = GetSetting("SourceCode", "Settings", "Icons")
    Trans = GetSetting("SourceCode", "Settings", "TransState")
    ListDock = GetSetting("SourceCode", "Settings", "AlignList")
    Select Case IconSize
        Case 0
            frmSettings.optLarge.Value = True
            frmSettings.optSmall.Value = False
        Case 1
            frmSettings.optSmall.Value = True
            frmSettings.optLarge.Value = False
    End Select
    Select Case Trans
        Case 0
            frmSettings.optTransOff.Value = True
            frmSettings.optTransOn.Value = False
        Case 1
            frmSettings.optTransOff.Value = False
            frmSettings.optTransOn.Value = True
    End Select
    Select Case ListDock
        Case 3
            frmSettings.optLeft.Value = True
            frmSettings.optRight.Value = False
        Case 4
            frmSettings.optLeft.Value = False
            frmSettings.optRight.Value = True
    End Select
    frmSettings.Show
End Sub

Private Sub mnuWindowArange_Click()
On Error Resume Next
    '============================================================
    ' This command agganges the minimized windows in an orderly
    ' fashion.
    '============================================================
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuHelpAbout_Click()
On Error Resume Next
    '============================================================
    ' This opens the About window.
    '============================================================
    frmAbout.Show
End Sub

Private Sub mnuMainAuth_Click()
On Error Resume Next
    '============================================================
    ' This opens the window to allow user to edit the Author's
    ' table.
    '============================================================
    frmEditAuth.Show
End Sub

Private Sub mnuMainLang_Click()
On Error Resume Next
    '============================================================
    ' This command opens the appropriate window to allow users to
    ' edit the Languages Table.
    '============================================================
    frmEditLang.Show
End Sub

Private Sub mnuWindowCascade_Click()
On Error Resume Next
    '============================================================
    ' This arranges the open windows in a cascaded layout.
    '============================================================
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileHoriz_Click()
On Error Resume Next
    '============================================================
    ' This tiles the open windows horizontally.
    '============================================================
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileVert_Click()
On Error Resume Next
    '============================================================
    ' This tiles the open windows vertically.
    '============================================================
    Me.Arrange vbTileHorizontal
End Sub

Private Sub tbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    '============================================================
    ' This handles the events sent when the user click on the
    ' toolbar.  Basically, it redirects the input based on the
    ' menus.
    '============================================================
    Select Case Button.Key
        Case "Open"
            mnuFileOpen_Click
        Case "Settings"
            mnuFileSettings_Click
        Case "Exit"
            mnuFileExit_Click
        Case "Add"
            mnuDataAdd_Click
        Case "List"
            cmdRefresh_Click
        Case "Find"
            mnuDataSearch_Click
        Case "Project"
            mnuDataProjList_Click
        Case "Source"
            mnuDataSourceList_Click
        Case "Author"
            mnuMainAuth_Click
        Case "Lang"
            mnuMainLang_Click
        Case "Cascade"
            mnuWindowCascade_Click
        Case "TileV"
            mnuWindowTileVert_Click
        Case "TileH"
            mnuWindowTileHoriz_Click
        Case "Icons"
            mnuWindowArange_Click
        Case "About"
            mnuHelpAbout_Click
    End Select
End Sub

Private Sub GetData(ListType As String)
On Error Resume Next
    '============================================================
    ' This is the sub that is responsible for populating the List
    ' and opening the correct table.  The first thing that
    ' happens is the database is opened, and then, based on the
    ' ListType, the recordset is established.  Then the list is
    ' cleared, populated and the records are counted and the
    ' status bar text is filled in with the current record count.
    '============================================================
    Dim Count As Integer
    Set db = OpenDatabase(DBPath)
    Select Case ListType
        Case "Source"
            Set rsCode = db.OpenRecordset("tblSourceCode", dbOpenDynaset)
        Case "Project"
            Set rsCode = db.OpenRecordset("tblProjectCode", dbOpenDynaset)
    End Select
    lstCode.Clear
    rsCode.MoveFirst
    Do While Not rsCode.EOF
        lstCode.AddItem (rsCode.Fields("CodeName"))
        Count = Count + 1
        rsCode.MoveNext
    Loop
    If Count > 1 Then
        stbInfo.Panels(1).Text = "There are currently " & Count & " code examples in the database."
    Else
        stbInfo.Panels(1).Text = "There is currently " & Count & " code example in the database."
    End If
End Sub

Function GetButtons()
    '============================================================
    ' Here the images that match up with the buttons on the
    ' toolbar are set.  It also checks the ListType to show the
    ' appropriate image for the current ListType.
    '============================================================
    ListType = GetSetting("SourceCode", "Settings", "DataSet")
    tbTools.Buttons.Item(1).Image = 1
    tbTools.Buttons.Item(2).Image = 2
    tbTools.Buttons.Item(3).Image = 3
    tbTools.Buttons.Item(5).Image = 4
    tbTools.Buttons.Item(6).Image = 5
    tbTools.Buttons.Item(7).Image = 6
    Select Case ListType
        Case "Source"
            tbTools.Buttons.Item(9).Image = 7
            tbTools.Buttons.Item(10).Image = 10
        Case "Project"
            tbTools.Buttons.Item(9).Image = 8
            tbTools.Buttons.Item(10).Image = 9
    End Select
    tbTools.Buttons.Item(12).Image = 11
    tbTools.Buttons.Item(13).Image = 12
    tbTools.Buttons.Item(15).Image = 13
    tbTools.Buttons.Item(16).Image = 14
    tbTools.Buttons.Item(17).Image = 15
    tbTools.Buttons.Item(18).Image = 16
    tbTools.Buttons.Item(20).Image = 17
    tbTools.Buttons.Item(21).Image = 18
End Function
