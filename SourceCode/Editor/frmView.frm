VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmView 
   Caption         =   "Animation Viewer"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9315
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5415
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15928
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      Begin VB.VScrollBar UDBar 
         Height          =   4695
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Outer 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   120
         ScaleHeight     =   4815
         ScaleWidth      =   2055
         TabIndex        =   3
         Top             =   240
         Width           =   2055
         Begin VB.PictureBox Inner 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4695
            Left            =   0
            ScaleHeight     =   4695
            ScaleMode       =   0  'User
            ScaleWidth      =   1979.587
            TabIndex        =   4
            Top             =   0
            Width           =   2055
            Begin VB.CommandButton cmdScene 
               Appearance      =   0  'Flat
               Caption         =   "Command1"
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Visible         =   0   'False
               Width           =   2055
            End
         End
      End
   End
   Begin VB.PictureBox view 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   2400
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Menu optFile 
      Caption         =   "File"
      Index           =   2
      Begin VB.Menu mnuFileOP 
         Caption         =   "Open"
         Index           =   1
      End
      Begin VB.Menu mnuFileOP 
         Caption         =   "Exit"
         Index           =   2
      End
   End
   Begin VB.Menu optHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpOP 
         Caption         =   "Help"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Start(FileName As String)
    For N = cmdScene.Count - 1 To 1 Step -1
    Unload cmdScene(N):    Next
    For N = 1 To 16
        Load cmdScene(N)
        cmdScene(N).Visible = True
        cmdScene(N).Top = ((N - 1) * cmdScene(N).Height) * 1.1
        cmdScene(N).Caption = N
    Next
    N = cmdScene.Count - 1
    If N <> 0 Then
        Inner.Height = ((N - 1) * cmdScene(N).Height) * 1.1
    End If
    If Inner.ScaleHeight < Outer.ScaleHeight Then
        UDBar.Visible = False
    Else
        UDBar.Max = Inner.ScaleHeight - Outer.Height + 2150
        UDBar.SmallChange = UDBar.Max / 20
        UDBar.LargeChange = UDBar.Max / 5
    End If
    If FileName <> "" Then
        LoadCompressedModel FileName
        Faster.RunEngine view
    End If
    If Inner.ScaleHeight > Outer.ScaleHeight Then
        UDBar.Visible = True
        For N = 1 To cmdScene.Count - 1
            cmdScene(N).Width = 1632.798
        Next N
      Else
        UDBar.Visible = False
        For N = 1 To cmdScene.Count - 1
            cmdScene(N).Width = 1979.587
        Next N
    End If
End Sub

Private Sub cmdScene_Click(Index As Integer)
    For N = 1 To Comp.SkelitonCount
        World.Morph(N).Angle.X = Rnd * 50
    Next N: Faster.RunEngine view
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    view.Height = Me.ScaleHeight - SBar.Height
    view.Width = Me.ScaleWidth - view.Left
    Frame1.Height = Me.ScaleHeight - SBar.Height
    Outer.Height = Me.ScaleHeight - 500 - SBar.Height
    UDBar.Height = Outer.Height:    UDBar.Top = Outer.Top
    UDBar.Max = Inner.ScaleHeight - Outer.Height + (Outer / 4)
    UDBar.SmallChange = Abs(UDBar.Max / 20)
    UDBar.LargeChange = Abs(UDBar.Max / 5)
    Xf = view.ScaleWidth / 2
    Yf = view.ScaleHeight / 2
    If Inner.ScaleHeight > Outer.ScaleHeight Then
        UDBar.Visible = True
        For N = 1 To cmdScene.Count - 1
            cmdScene(N).Width = 1632.798
        Next N
      Else
        UDBar.Visible = False
        For N = 1 To cmdScene.Count - 1
            cmdScene(N).Width = 1979.587
        Next N
    End If
    Faster.RunEngine view
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Comp.FaceCount = 0: Comp.VertexCount = 0
    Comp.SkelitonCount = 0: Comp.WeaponCount = 0
    frmView.Visible = False: frmMain.Visible = True
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub

Private Sub mnuFileOP_Click(Index As Integer)
    Select Case Index
    Case 1
        FileName = SelectFileName("Compile", "Select compiled file to view")
        If frmMain.GetFile.FileName <> "" Then
            If Faster.LoadCompressedModel(frmMain.GetFile.FileName) = False Then
                MsgBox "Couldn't load this file, make sure its the one you want!", , "Error"
            End If
            Faster.RunEngine view
        End If
    Case 2
        Unload Me
    End Select
End Sub

Private Sub UDBar_Change()
    Inner.Top = -UDBar
End Sub

Private Sub UDBar_Scroll()
    Inner.Top = -UDBar
End Sub

Private Sub view_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static oX As Integer, oY As Integer
    If Button = 1 Then
        World.Angle.X = (World.Angle.X - (Y - oY)) Mod 360
        World.Angle.Y = (World.Angle.Y - (X - oX)) Mod 360
        Faster.RunEngine view: Refresh
    End If
    oX = X: oY = Y
End Sub
