VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Shop 6 Settings"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "FrmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3135
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5895
      Begin MSComctlLib.Slider sldGrey 
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   2280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Max             =   25
         TickFrequency   =   5
      End
      Begin VB.PictureBox Clours 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   705
         TabIndex        =   15
         ToolTipText     =   "A line used as a center point on some shapes"
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox Clours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   705
         TabIndex        =   9
         ToolTipText     =   "A dotted line around each selected item"
         Top             =   1080
         Width           =   735
      End
      Begin VB.PictureBox Clours 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4800
         ScaleHeight     =   225
         ScaleWidth      =   705
         TabIndex        =   8
         ToolTipText     =   "The guidelines of a new shape that you can't alter"
         Top             =   1080
         Width           =   735
      End
      Begin VB.PictureBox Clours 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4800
         ScaleHeight     =   225
         ScaleWidth      =   705
         TabIndex        =   7
         ToolTipText     =   "The guidelines of a new shape that you can alter"
         Top             =   600
         Width           =   735
      End
      Begin VB.PictureBox Clours 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   705
         TabIndex        =   6
         ToolTipText     =   "A box surrounding all of the selected items"
         Top             =   600
         Width           =   735
      End
      Begin MSComDlg.CommonDialog GetColour 
         Left            =   5040
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label lblgx 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grey out shade"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Object outline"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Reflected obejct guideline"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "New object guideline"
         Height          =   195
         Left            =   3120
         TabIndex        =   12
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Box Guide"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Axis"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2895
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CheckBox chkCenter 
         Caption         =   "Always center when changing views"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "If on, then the view centers on the selected object when you switch views"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox GridSize 
         Height          =   285
         Left            =   1095
         TabIndex        =   17
         Text            =   "10"
         Top             =   2280
         Width           =   615
      End
      Begin VB.CheckBox chkHilighBox 
         Caption         =   "Highlight corners and edges of selected items box"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Highligts the corners and middles of the selection box"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CheckBox chkNew 
         Caption         =   "Ask for a model name when you press 'new'"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Automaticly give a new model a name"
         Top             =   240
         Width           =   4095
      End
      Begin VB.CheckBox DataSpin 
         Caption         =   "Display UpDown control on data grid - May slow down old machines"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Used on the grid in the animation editor"
         Top             =   720
         Value           =   1  'Checked
         Width           =   5415
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1711
         TabIndex        =   16
         ToolTipText     =   "Click this to change the grid size"
         Top             =   2280
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "GridSize"
         BuddyDispid     =   196611
         OrigLeft        =   2400
         OrigTop         =   960
         OrigRight       =   2640
         OrigBottom      =   1215
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label8 
         Caption         =   "Grid Size"
         Height          =   255
         Left            =   255
         TabIndex        =   18
         Top             =   2280
         Width           =   735
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Colours"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Clours_Click(Index As Integer)
    On Error GoTo PressedCancel
        GetColour.DialogTitle = "Select a new colour"
        GetColour.ShowColor
        Clours(Index).BackColor = GetColour.Color
        Colours(Index + 3) = GetColour.Color
        If Index = 4 Then
             Colours(Index + 3) = 65536 - GetColour.Color
        End If
        frmMain.Axis.BorderColor = Colours(Index + 3)
PressedCancel:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    Cancel = 1
    If IsThisValid(GridSize, 1) = False Then
        GridSize = UpDown1
    End If
    Me.Visible = False
    frmMain.DrawModel
    DrawGuide
End Sub

Private Sub sldGrey_Click()
    lblgx.BackColor = RGB(sldGrey * 10, sldGrey * 10, sldGrey * 10)
End Sub

Private Sub TabStrip1_Click()
    For N = 0 To Frame.Count - 1: Frame(N).Visible = False: Next N
    Frame(TabStrip1.SelectedItem.Tag - 1).Visible = True
End Sub
