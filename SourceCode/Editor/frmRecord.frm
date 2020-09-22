VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record Rotation"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   ControlBox      =   0   'False
   Icon            =   "frmRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5400
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Okay"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox Dire 
      Caption         =   "Rotate on Z axis"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox Dire 
      Caption         =   "Rotate on Y axis"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.CheckBox Dire 
      Caption         =   "Rotate on X axis"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox lblFileName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2880
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog GetFile 
      Left            =   2520
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.UpDown SetAngle 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   2040
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Angle"
      BuddyDispid     =   196619
      OrigLeft        =   1680
      OrigTop         =   2040
      OrigRight       =   1920
      OrigBottom      =   2295
      Max             =   360
      Min             =   2
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Angle 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Text            =   "36"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton SetFile 
      Cancel          =   -1  'True
      Caption         =   "..."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   255
   End
   Begin VB.ListBox Rotations 
      Height          =   1230
      ItemData        =   "frmRecord.frx":030A
      Left            =   3600
      List            =   "frmRecord.frx":032C
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "How many &frames per rotation do you want?"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Which way do you want to spin the model?"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "&Save to:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "How many &times do you want it to spin?"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cancel = 1
    If lblFileName = "" Then MsgBox "Enter a filename first!", 48, "Error": Exit Sub
    If IsThisValid(Angle, 0) = False Then MsgBox "The number of frames per rotation is an invalid value", 48, "Error": Exit Sub
    If Dire(0) = 0 And Dire(1) = 0 And Dire(2) = 0 Then MsgBox "You must tick atleast one axis", 48, "Error": Exit Sub
    Visible = False
    Me.Tag = "Run"
End Sub

Private Sub Command2_Click()
    Visible = False
    Me.Tag = "Cancel"
End Sub

Private Sub Form_Load()
    Rotations.Selected(0) = True
End Sub


Private Sub SetFile_Click()
    On Error GoTo NoSaveeither
    GetFile.DialogTitle = "Enter a file name to save to..."
    GetFile.Filter = "Bitmap image (*.bmp) |*.bmp"
    GetFile.ShowSave
    lblFileName = GetFile.FileName
    lblFileName.ToolTipText = GetFile.FileName
NoSaveeither:
End Sub
