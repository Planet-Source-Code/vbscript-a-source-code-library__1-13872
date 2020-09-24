VERSION 5.00
Begin VB.Form frmList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Current Source Code"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstCode 
      Columns         =   2
      Height          =   3960
      ItemData        =   "frmList.frx":0000
      Left            =   120
      List            =   "frmList.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "This pane displays the code title."
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh from Database"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Re-read the database information into the list."
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Code"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Edit highlighted code."
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View Code"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "View highlighted code."
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim db As Database
    Dim rsCode As Recordset

Public Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Icon = LoadResPicture(107, 1)
    cmdRefresh_Click
    DoGradient Me
End Sub

Public Sub cmdRefresh_Click()
End Sub

