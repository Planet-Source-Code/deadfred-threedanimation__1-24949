VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmObjectPrp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "frmObjectPrp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7020
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog GetColour 
      Left            =   5760
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame 
      Height          =   4575
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6615
      Begin VB.ListBox lstJoints 
         Height          =   2400
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   375
      End
      Begin VB.PictureBox View 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   413
         TabIndex        =   3
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label BackColour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Colour"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Object Targeted to"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   3240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame 
      Height          =   4575
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cdmMore 
         Caption         =   "UnGroup"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   16
         ToolTipText     =   "Break apart any existing group"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton cdmMore 
         Caption         =   "Group"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   15
         ToolTipText     =   "Group the selected objects together"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblFace 
         Caption         =   "Unavalible"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lvlVertex 
         Caption         =   "Unavalible"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblCount 
         Caption         =   "Unavalible"
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of objects selected"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Faces in selected objects"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Vertecies in selected objects"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
      End
   End
   Begin MSComctlLib.TabStrip Tabs 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8916
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "All Objects"
            Object.Tag             =   "0"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Statistics"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmObjectPrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xofs As Integer, Yofs As Integer
Dim Angle1, Angle2, Angle3
Dim Xmouse As Integer, Ymouse As Integer
Dim Xcen As Integer, Ycen As Integer, Zcen As Integer

Public Sub RunAtStart()
    X = Functions.CountSelectedObject
    If X = 1 Then
        Me.Caption = "One object selected"
        Label1 = "Object Targeted to : " & Object(ObjectSelected).Vertex(1).TargetName
      Else
        Caption = X & " objects selected"
    End If
    Xofs = view.ScaleWidth / 2
    Yofs = view.ScaleHeight / 2
    Dim N As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Selected = True And Object(N).Used = True Then
            Xcen = Xcen + FindCenter("X", N)
            Ycen = Ycen + FindCenter("Y", N)
            Zcen = Zcen + FindCenter("Z", N)
        End If
    Next N
    Xcen = Xcen / CountSelectedObject
    Ycen = Ycen / CountSelectedObject
    Zcen = Zcen / CountSelectedObject
    
    Me.lblCount = Functions.CountSelectedObject
    Me.lblFace = Functions.CountSelectedFaces
    Me.lvlVertex = Functions.CountSelectedVertecies
    

     frmMain.GetFile.Color = Object(FirstObject).Colour
     BackColour.BackColor = frmMain.GetFile.Color

    frmJointProp.RunAtStart
    
    DrawME
End Sub



Private Sub cdmMore_Click(Index As Integer)
    Select Case Index
        Case 0: GroupSelected
        Case 1: UnGroupSelected
    End Select
End Sub

Private Sub Command1_Click()
    lstJoints.Visible = True
    For N = 1 To 10
        Refresh
        lstJoints.Height = 240 * N
        lstJoints.Top = Command1.Top - lstJoints.Height
    Next N
    lstJoints.SetFocus
End Sub

Private Sub Command2_Click()
    GetColour.ShowColor
    BackColour.BackColor = GetColour.Color
    For N = 1 To cstTotalObjects
        If Object(N).Selected = True Then
            Object(N).Colour = GetColour.Color
        End If
    Next N
    DrawME
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    frmMain.SetFocus
    frmMain.DrawModel
End Sub

Private Sub Frame_Click(Index As Integer)
    lstJoints.Visible = False
End Sub

Private Sub lstJoints_LostFocus()
    lstJoints.Visible = False
End Sub

Private Sub Tabs_Click()
    lstJoints.Visible = False
    For N = 0 To Frame.Count - 1: Frame(N).Visible = False: Next N
    Frame(Tabs.SelectedItem.Tag).Visible = True
End Sub

Private Sub DrawME()
    view.Cls
    For N = 1 To cstTotalObjects
        If Object(N).Used = True And Object(N).Selected = True Then
            ThreeDEngine.Draw3DBrush view, N, Angle1, Angle2, Angle3, 0, 0, 0, Xcen, Ycen, Zcen
        End If
    Next N
    view.Refresh
End Sub

Private Sub View_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Angle2 = Angle2 + (Xmouse - X)
        Angle1 = Angle1 + (Ymouse - Y)
        DrawME
    End If
    If Button = 2 Then
        Xofs = Xofs - (Xmouse - X)
        Yofs = Yofs - (Ymouse - Y)
        DrawME
    End If
    Xmouse = X: Ymouse = Y
End Sub
