VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form SceneEdit 
   Caption         =   "Scene Editor"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8910
   Icon            =   "SceneEdit.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList Icons 
      Left            =   720
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SceneEdit.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SceneEdit.frx":041E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SceneEdit.frx":073A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SceneEdit.frx":0A56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   4050
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7144
      BandBorders     =   0   'False
      _CBWidth        =   8175
      _CBHeight       =   4050
      _Version        =   "6.0.8169"
      Child1          =   "Frame"
      MinHeight1      =   1995
      Width1          =   2490
      NewRow1         =   0   'False
      Child2          =   "Main"
      MinHeight2      =   1995
      Width2          =   3810
      NewRow2         =   0   'False
      Child3          =   "View"
      MinHeight3      =   1995
      Width3          =   615
      NewRow3         =   -1  'True
      Begin VB.PictureBox View 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1995
         Left            =   165
         ScaleHeight     =   129
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   524
         TabIndex        =   15
         Top             =   2025
         Width           =   7920
         Begin VB.Label ShowName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label3"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2040
            TabIndex        =   16
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.PictureBox Main 
         BackColor       =   &H00FFFFFF&
         Height          =   1995
         Left            =   2655
         ScaleHeight     =   1935
         ScaleWidth      =   5370
         TabIndex        =   4
         Top             =   30
         Width           =   5430
         Begin VB.VScrollBar hsbSide 
            Height          =   4335
            LargeChange     =   10
            Left            =   5040
            TabIndex        =   7
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar vsbLow 
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   4560
            Width           =   975
         End
         Begin VB.CommandButton Box 
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   5
            Top             =   4320
            Width           =   255
         End
         Begin VB.PictureBox Inner 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   0
            ScaleHeight     =   2535
            ScaleWidth      =   4815
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   4815
            Begin VB.TextBox Data 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   240
               TabIndex        =   9
               Top             =   480
               Visible         =   0   'False
               Width           =   495
            End
            Begin MSComCtl2.UpDown Spin 
               Height          =   255
               Left            =   720
               TabIndex        =   10
               Top             =   480
               Visible         =   0   'False
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   450
               _Version        =   393216
               Increment       =   5
               Max             =   360
               Wrap            =   -1  'True
               Enabled         =   -1  'True
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Scale"
               Height          =   255
               Left            =   4080
               TabIndex        =   17
               Top             =   0
               Width           =   615
            End
            Begin VB.Line Line3 
               X1              =   3720
               X2              =   5040
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Label Head 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label1"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Number 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2"
               Height          =   195
               Index           =   0
               Left            =   75
               TabIndex        =   13
               Top             =   480
               Visible         =   0   'False
               Width           =   105
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Angle"
               Height          =   255
               Left            =   1200
               TabIndex        =   12
               Top             =   0
               Width           =   495
            End
            Begin VB.Line Line1 
               X1              =   840
               X2              =   2160
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line2 
               X1              =   2280
               X2              =   3600
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
               Height          =   255
               Left            =   2640
               TabIndex        =   11
               Top             =   0
               Width           =   615
            End
         End
      End
      Begin MSComctlLib.TreeView Frame 
         Height          =   1995
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3519
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   706
         LabelEdit       =   1
         PathSeparator   =   "_"
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "Icons"
         Appearance      =   1
         OLEDragMode     =   1
      End
   End
   Begin MSComctlLib.StatusBar Sbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Tag             =   "0"
      Top             =   5940
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "13:51"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10081
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Play mode off"
            TextSave        =   "Play mode off"
            Key             =   "OnOff"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   4560
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuMor 
      Caption         =   "&Options"
      Begin VB.Menu More 
         Caption         =   "&Play"
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu More 
         Caption         =   "&Stop"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   {F2}
      End
      Begin VB.Menu More 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu More 
         Caption         =   "Frames.."
         Index           =   4
         Shortcut        =   +{INSERT}
         Visible         =   0   'False
      End
      Begin VB.Menu More 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu More 
         Caption         =   "&Close"
         Index           =   6
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Display"
      Begin VB.Menu Options 
         Caption         =   "&Layout"
         Index           =   0
         Begin VB.Menu opLayout 
            Caption         =   "&Vertical"
            Index           =   1
            Shortcut        =   {F9}
         End
         Begin VB.Menu opLayout 
            Caption         =   "&Horizontal"
            Index           =   2
            Shortcut        =   {F11}
         End
         Begin VB.Menu opLayout 
            Caption         =   "&Small View"
            Index           =   3
            Shortcut        =   {F12}
         End
      End
      Begin VB.Menu Options 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Options 
         Caption         =   "&Show Skeliton"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu Options 
         Caption         =   "Show &Model"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu Options 
         Caption         =   "Show &Origin"
         Checked         =   -1  'True
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu Help 
         Caption         =   "&Help"
         Index           =   1
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu PopUpEdit 
      Caption         =   "PopUpEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Index           =   2
      End
   End
   Begin VB.Menu Popop 
      Caption         =   "Popop"
      Visible         =   0   'False
      Begin VB.Menu popup 
         Caption         =   "Rename"
         Index           =   1
      End
      Begin VB.Menu popup 
         Caption         =   "Delete"
         Index           =   2
      End
      Begin VB.Menu popup 
         Caption         =   "Remove"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu popup 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu popup 
         Caption         =   "Add 1"
         Index           =   5
      End
      Begin VB.Menu popup 
         Caption         =   "Add 5"
         Index           =   6
      End
      Begin VB.Menu popup 
         Caption         =   "Add Increment frame"
         Index           =   7
      End
   End
End
Attribute VB_Name = "sceneEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sceneon As Byte, FrameOn As Integer, JointOn As Byte
Dim Xofs As Integer, Yofs As Integer
Dim Angle1, Angle2, Angle3
Dim Xmouse As Integer, Ymouse As Integer
Dim Xstart As Integer, Ystart As Integer
Dim HaHa2D(cstTotalJoints, 2) As Integer
Dim MakeWork(cstTotalJoints) As Coord
Dim CopySpace(cstTotalJoints, 6) As Integer

Sub Start()
    Inner.Visible = False
    FrameOn = 1: JointOn = 1
    Xofs = view.ScaleWidth / 2
    Yofs = view.ScaleHeight / 2
    Num = 0
    For N = 1 To cstTotalJoints
        KeepRight(N, 1) = 0
    Next N
    SetUpGrid (CountJoints)
    Me.Visible = True
    Me.SetFocus
    DoEvents
    sceneon = 1
    DrawME
End Sub

Private Sub CoolBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawWindowsSize
End Sub

Private Sub Frame_NodeClick(ByVal Node As MSComctlLib.Node)
    If Frame.SelectedItem.Tag = 4 Then
        Label3.Caption = "Working file - " & Frame.SelectedItem.Key
        SBar.Panels(2).Text = "Increment Scene - " & Frame.SelectedItem.Key & "    (" & Frame.SelectedItem.Text & ")"
        Inner.Visible = True
    End If
    If Frame.SelectedItem.Tag = 3 Then
        Label3.Caption = "Working file - " & Frame.SelectedItem.Key
        SBar.Panels(2).Text = "Scene - " & Frame.SelectedItem.Key & "    (" & Frame.SelectedItem.Text & ")"
        Inner.Visible = True
    End If
    If Frame.SelectedItem.Tag = 2 Then
        Label3.Caption = "<"
        SBar.Panels(2).Text = "Scene - " & Frame.SelectedItem.Key & "    (" & Frame.SelectedItem.Text & ")"
        Inner.Visible = False
    End If
    If Frame.SelectedItem.Tag = 1 Then
        Label3.Caption = "<"
        SBar.Panels(2).Text = "Model Name - " & Frame.SelectedItem.Text
        Inner.Visible = False
    End If
End Sub

Private Sub Help_Click(Index As Integer)
    ShowHelp "Creating scenes"
End Sub

Private Sub Main_Paint()
'    UpdateWindowz
'    DrawME
End Sub

Private Sub More_Click(Index As Integer)
    If Index = 1 Then
        More(1).Checked = True
        More(2).Checked = False
        SBar.Tag = 1: SBar.Panels(3).Text = "Play mode on"
    End If
    If Index = 2 Then
        More(2).Checked = True
        More(1).Checked = False
        SBar.Tag = 0: SBar.Panels(3).Text = "Play mode off"
    End If
    If Index = 4 Then
        If Frame.SelectedItem.Tag = 1 Then
            MsgBox "You must select a scene from the list first"
            Exit Sub
        End If
        
        Me.Enabled = False
        frmAddframe.Visible = True
        Do: DoEvents
        Loop While frmAddframe.Visible = True
    
        If Frame.SelectedItem.Tag = 3 Then
            Frame.SelectedItem.Parent.Selected = True
        End If
    
        MsgBox "Insert frames here" & vbNewLine & Frame.SelectedItem.Text
        NewChildren = Frame.SelectedItem.Children
        Frame.SelectedItem.Child.Selected = True
        
        If frmAddframe.optDuplicate = True Then
        
            For N = 1 To NewChildren
        
                
                
                Me.Add_Frame
                MsgBox ""
                'Frame.SelectedItem.Next.Selected = True
                Refresh
            Next N
        
        End If
    
    End If
    If Index = 6 Then
        Label3 = "<###>"
        Me.Visible = False: frmMain.Visible = True
        More(2).Checked = True
        UpdateToolbarlist
        frmMain.SetFocus
        Label3.Tag = ""
    End If
End Sub

Private Sub opLayout_Click(Index As Integer)
    Select Case Index
        Case 1
            CoolBar1.Bands(1).MinHeight = (Me.ScaleHeight - SBar.Height)
            CoolBar1.Bands(1).Width = Me.ScaleWidth / 3
            CoolBar1.Bands(2).MinHeight = (Me.ScaleHeight - SBar.Height)
            CoolBar1.Bands(2).NewRow = False
            CoolBar1.Bands(2).Width = Me.ScaleWidth / 3
            CoolBar1.Bands(3).MinHeight = (Me.ScaleHeight - SBar.Height)
            CoolBar1.Bands(3).NewRow = False
            CoolBar1.Bands(3).Width = Me.ScaleWidth / 3
        Case 2
            CoolBar1.Bands(1).MinHeight = (Me.ScaleHeight - SBar.Height) / 3
            CoolBar1.Bands(2).NewRow = True
            CoolBar1.Bands(2).Width = Me.ScaleWidth
            CoolBar1.Bands(2).MinHeight = (Me.ScaleHeight - SBar.Height) / 3
            CoolBar1.Bands(3).NewRow = True
            CoolBar1.Bands(3).Width = Me.ScaleWidth
            CoolBar1.Bands(3).MinHeight = (Me.ScaleHeight - SBar.Height) / 3
        Case 3
            CoolBar1.Bands(1).MinHeight = (Me.ScaleHeight - SBar.Height) / 2
            CoolBar1.Bands(1).Width = Me.ScaleWidth / 5
            CoolBar1.Bands(2).NewRow = False
            CoolBar1.Bands(2).Width = (Me.ScaleWidth / 5) * 4
            CoolBar1.Bands(2).MinHeight = (Me.ScaleHeight - SBar.Height) / 2
            CoolBar1.Bands(3).NewRow = True
            CoolBar1.Bands(3).Width = Me.ScaleWidth
            CoolBar1.Bands(3).MinHeight = (Me.ScaleHeight - SBar.Height) / 2
    End Select
End Sub

Private Sub Options_Click(Index As Integer)
    If Index > 1 Then
        If Options(Index).Checked = True Then
            Options(Index).Checked = False
        Else
            Options(Index).Checked = True
        End If
        DrawME
    End If
End Sub

Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For N = 1 To 100
        If BaseFrame(N).Used = True Then
            BaseFrame(N).Selected = False
            XX = HaHa(N, 1)
            YY = HaHa(N, 2)
            If Almostt(X, Y, XX, YY, 4) = True Then
                JointOn = N
                BaseFrame(N).Selected = True
                view.ToolTipText = BaseFrame(N).Name
            End If
        End If
    Next N
    For N = 1 To Number.Count - 1
        Number(N).FontBold = False
    Next N
    Number(KeepRight(JointOn, 2)).FontBold = True
    DrawME
    Spin.Visible = False
End Sub

Private Sub View_Paint()
    DrawME
End Sub

Private Sub vsbLow_Change()
    Inner.Left = vsbLow
End Sub

Private Sub vsbLow_Scroll()
    Inner.Left = vsbLow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Label3 = "<###>"
    Me.Visible = False: frmMain.Visible = True
    More(2).Checked = True
    frmMain.Visible = True
    frmMain.SetFocus: Cancel = 1
    UpdateToolbarlist
    Label3.Tag = ""
End Sub

Private Sub Inner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu PopUpEdit
End Sub

Private Sub Label3_Change()
    On Error GoTo BustedBoyo
    If Data.Count = 1 Then Exit Sub
    Close
    If sceneEdit.Visible = True Then
        SBar.Panels(2) = Label3
        If Label3.Tag <> "" Then
            Open Label3.Tag For Output As #1
            For M = 1 To CountJoints
                Print #1, "Joint ,"; M;
                For N = 1 To 9
                    Print #1, ","; Data((N + ((M - 1) * 9)));
                Next N
                Print #1,
            Next M
            Print #1, "End"
            Close
        End If
    End If
    Close
    If Mid(Label3, 1, 1) = "<" Then Exit Sub
    workingfile = App.Path + "\frames\" & Frame.SelectedItem.Key & ".dat"
    Open workingfile For Input As #1
 
Fool2:
    Input #1, temp
    If temp = "Joint" Then
        Input #1, M
        For N = 1 To 9
            Input #1, valu
            Data(N + ((M - 1) * 9)) = valu
        Next N
        GoTo Fool2:
    End If
    Close
    Label3.Tag = workingfile
    DrawME
BustedBoyo:
End Sub

Private Sub Main_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu PopUpEdit
End Sub


Private Sub mnuEdit_Click(Index As Integer)
 Select Case Index
  Case 1
   For N = 1 To CountJoints
    For M = 1 To 6
     CopySpace(N, M) = Data(((N - 1) * 6) + M)
     rdo = rdo + Str(CopySpace(N, M)) & " "
    Next M
    rdo = rdo + vbNewLine
   Next N
  Case 2
   For N = 1 To CountJoints
    For M = 1 To 6
     Data(((N - 1) * 6) + M) = CopySpace(N, M)
    Next M
   Next N
 End Select
End Sub

Private Sub Number_Click(Index As Integer)
    JointOn = Index
    For N = 1 To Number.Count - 1
        Number(N).FontBold = False
    Next N
    Number(JointOn).FontBold = True
    For N = 1 To cstTotalJoints
        BaseFrame(N).Selected = False
    Next N
    BaseFrame(JointOn).Selected = True
    DrawME
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

Private Sub DrawME()
    view.Cls
    FindNewSkeliton
    If Options(4).Checked = True Then DrawGuides view, Angle1, Angle2, Angle3, view.ScaleWidth / 2, view.ScaleHeight / 2
    If Options(3).Checked = True Then
        For N = 1 To cstTotalObjects
            If Object(N).Used = True Then
                ThreeDEngine.Draw3DBrush view, N, Angle1, Angle2, Angle3, 0, 0, 0, 0, 0, 0
            End If
        Next N
    End If
    If Options(2).Checked = True Then Draw3DSkeliton view, Angle1, Angle2, Angle3, view.ScaleWidth / 2, view.ScaleHeight / 2
    view.Refresh
End Sub

Sub SetUpGrid(Roze)
    If Roze = 0 Then Exit Sub
    
    SetUPSkeliton
    
    Dim QuitThis As Boolean
    If (Number.Count - 1) = Roze Then
        QuitThis = True
    End If
    If Number.Count > 1 Then
        For N = 1 To Number.Count - 1
            Unload Number(N)
        Next N
    End If
    OldWide = 0
    For N = 1 To Roze
        X = X + 1
        Load Number(X)
        Number(X).Visible = True
        Number(X) = N
        Number(X).Top = ((Data(0).Height - 17) * N) + 300
        Number(N).Caption = BaseFrame(KeepRight(N, 1)).Name
        Number(X).Left = 100
        Number(X).ToolTipText = BaseFrame(X).Name
        If Number(N).Width > OldWide Then OldWide = Number(N).Width
    Next N
    OldWide = OldWide
    Label1.Left = OldWide + 800
    Line1.x1 = OldWide + 500
    Line1.x2 = OldWide + 1700
    Label2.Left = OldWide + 2300
    Line2.x1 = OldWide + 2000
    Line2.x2 = OldWide + 3200
    
    Label4.Left = OldWide + 3800
    Line3.x1 = OldWide + 3500
    Line3.x2 = OldWide + 4600
    X = 0
    If CountJoints > 25 Then
        Timei.Visible = True
        Timei.Caption = "Generating grid - Please wait..."
        Timei.TimeLeft.ToolTipText = ""
        Timei.TimeLeft.Max = Roze
        Timei.TimeLeft.Value = 0
    End If
    If Data.Count > 1 Then
        For N = 1 To Data.Count - 1
            Unload Data(N)
        Next N
    End If
    If Head.Count > 1 Then
        For N = 1 To Head.Count - 1
            Unload Head(N)
        Next N
    End If
    MLeft = 400
    X = 0
    For N = 0 To 8
    X = X + 1
    Load Head(X)
    If X = 1 Then Head(X).Caption = "X"
    If X = 2 Then Head(X).Caption = "Y"
    If X = 3 Then Head(X).Caption = "Z"
    If X = 4 Then Head(X).Caption = "X"
    If X = 5 Then Head(X).Caption = "Y"
    If X = 6 Then Head(X).Caption = "Z"
    If X = 7 Then Head(X).Caption = "X"
    If X = 8 Then Head(X).Caption = "Y"
    If X = 9 Then Head(X).Caption = "Z"
    Head(X).Visible = True
    Head(X).Left = OldWide + MLeft + (Head(X).Width - 17) * N
    Head(X).Top = 300
    Next N
    X = 0
    For N = 1 To Roze
        Timei.TimeLeft.Value = N
        Timei.TimeLeft.Refresh
        For M = 0 To 8
            X = X + 1
            Load Data(X)
            Data(X).Visible = True
            Data(X).Left = OldWide + MLeft + (Data(X).Width - 17) * M
            Data(X).Top = ((Data(X).Height - 17) * N) + 300
            Data(X).Text = "0"
            ' Data(X).ToolTipText = BaseFrame(KeepRight(n, 1)).Name
        Next M
    Next N
    Inner.Width = OldWide + 300 + MLeft + (Data(0).Width - 16) * 9
    Inner.Height = 500 + (Data(0).Height - 16) * (Roze + 1)
    UpdateWindowz
    Timei.Visible = False
End Sub

Sub UpdateWindowz()
    On Error GoTo ToSmall
    CoolBar1.Width = Me.ScaleWidth
    hsbSide.Left = Main.ScaleWidth - hsbSide.Width
    hsbSide.Height = Main.ScaleHeight - vsbLow.Height
    vsbLow.Top = Main.ScaleHeight - vsbLow.Height
    vsbLow.Width = Main.ScaleWidth - hsbSide.Width
    hsbSide.Max = Main.ScaleHeight - Inner.ScaleHeight - vsbLow.Height
    vsbLow.Max = Main.ScaleWidth - Inner.ScaleWidth - hsbSide.Width
    If hsbSide.Max <> 0 Then hsbSide.LargeChange = Abs(hsbSide.Max) / 4
    If vsbLow.Max <> 0 Then vsbLow.LargeChange = Abs(vsbLow.Max / 4)
    Box.Top = vsbLow.Top: Box.Left = hsbSide.Left
    Box.Height = vsbLow.Height: Box.Width = hsbSide.Width
    hsbSide.Enabled = False
    vsbLow.Enabled = False
    If Inner.Width > Main.ScaleWidth Then vsbLow.Enabled = True
    If Inner.Height > Main.ScaleHeight Then hsbSide.Enabled = True
ToSmall:
End Sub

Private Sub Data_GotFocus(Index As Integer)
    Model.Saved = False
    JointOn = Int((Index - 1) / 9) + 1
    For N = 1 To Number.Count - 1
        Number(N).FontBold = False
    Next N
    For N = 1 To 100: BaseFrame(N).Selected = False: Next N
    Number(JointOn).FontBold = True
    BaseFrame(KeepRight(JointOn, 1)).Selected = True
    DrawME
    Spin.Visible = False
    If frmSettings.DataSpin = 0 Then Exit Sub
    Spin.Tag = Index
    Spin.Visible = True
    colm = Index Mod 9: If colm = 0 Then colm = 9
    If colm > 2 Then Spin.Max = 1000: Spin.Min = -1000
    If colm < 4 Then Spin.Max = 355: Spin.Min = 0
    Spin.Left = Data(Index).Left + Data(Index).Width - 17
    Spin.Top = Data(Index).Top
    Spin.Height = Data(Index).Height
    On Error Resume Next
    Spin.Value = Data(Index)
End Sub

Private Sub Data_LostFocus(Index As Integer)
    Spin.Visible = False
    On Error GoTo NotNumber
    If Data(Index) = "" Then Exit Sub
    X = Data(Index) / 2
    Exit Sub
NotNumber:
    Beep
    Data(Index).SetFocus
End Sub

Private Sub Form_Load()
    UpdateWindowz
    frmMain.Tag = "Not Saved"
End Sub

Private Sub hsbSide_Change()
    Inner.Top = hsbSide
End Sub

Private Sub hsbSide_Scroll()
    Inner.Top = hsbSide
End Sub

Private Sub Spin_Change()
    Data(Spin.Tag) = Spin.Value
    DrawME
End Sub

Private Sub Form_Resize()
    DrawWindowsSize
End Sub

Private Sub DrawWindowsSize()
    On Error GoTo ToSmaall
    Row = 1
    CoolBar1.Width = Me.ScaleWidth
    If CoolBar1.Bands(2).NewRow = True Then Row = Row + 1
    If CoolBar1.Bands(3).NewRow = True Then Row = Row + 1
    CoolBar1.Bands(1).MinHeight = (Me.ScaleHeight - SBar.Height - 50) / Row
    CoolBar1.Bands(2).MinHeight = (Me.ScaleHeight - SBar.Height - 50) / Row
    CoolBar1.Bands(3).MinHeight = (Me.ScaleHeight - SBar.Height - 50) / Row
    If Tabz.Tag = 1 Then view.Left = Frame.Left
    If Tabz.Tag = 2 Then view.Left = Main.Left
    If Tabz.Tag = 3 Then view.Left = Main.Left + Main.Width + 100
    Tabz.Width = Me.ScaleWidth
    Tabz.Height = Me.ScaleHeight - SBar.Height
    Frame1.Width = Me.ScaleWidth - Frame1.Left - 100
    Frame1.Height = Me.ScaleHeight - Frame1.Top - 100 - SBar.Height
    view.Width = Frame1.Width - view.Left
    Xofs = view.ScaleWidth / 2
    Yofs = view.ScaleHeight / 2
    Frame.Height = Frame1.Height - Frame.Top - 100
    Main.Height = Frame.Height
    view.Height = Frame.Height
ToSmaall:
    UpdateWindowz
    DrawME
End Sub

Public Sub Add_Frame()
    Frame.Nodes(Frame.SelectedItem.Key).Expanded = True
    Scn = Val(Mid(Frame.SelectedItem.Key, 6, 3))
    If Frame.SelectedItem.Children = 0 Then
        Xxx = 1
    Else
        Yyy = Frame.Nodes(Frame.Nodes(Frame.SelectedItem.Key).Child.Key).LastSibling.Key
        Xxx = Val(Mid(Yyy, 8, 2)) + 1
    End If
    Namez = Frame.SelectedItem.Key & "_" & Xxx
    Frame.Nodes.Add Frame.SelectedItem.Key, 4, Namez, "Frame" & Xxx, 1
    Frame.Nodes(Namez).Tag = 3
    NewFile = App.Path + "\frames\" & Namez & ".dat"
    Open NewFile For Output As #1
    For M = 1 To CountJoints
        Print #1, "Joint ,"; M;
        For N = 1 To 9
            Print #1, ","; 0;
        Next N
        Print #1,
    Next M
    Print #1, "End"
    Close
End Sub

Function FindScene()
    For N = 1 To cstTotalScenes
        If Scenes(N).Used = False Then
            FindScene = N: Exit Function
        End If
    Next N
    FindScene = 0
End Function

Public Sub Add_Scene(NewName)
    Frame.Nodes("Model").Expanded = True
    Add = FindScene
    If FindScene = 0 Then
        MsgBox "Sorry, you can only have " & cstTotalScenes & " frames": Exit Sub
    End If
    Key = "Scene" & Add
    Frame.Nodes.Add "Model", 4, Key, NewName, 2
    Frame.Nodes(Key).Tag = 2
    Scenes(Add).Key = Key
    Scenes(Add).Name = NewName
    Scenes(Add).Used = True
    Scenes(Add).Mode = 1
    UpdateToolbarlist
End Sub

Public Sub UpdateToolbarlist()
    frmMain.MnuScene(0).Visible = True
    For N = frmMain.MnuScene.Count - 1 To 1 Step -1
        Unload frmMain.MnuScene(N)
    Next N
    nn = 0
    For N = 1 To sceneEdit.Frame.Nodes.Count
        XX = Mid(Frame.Nodes.Item(N).Key, 1, 5)
        If XX = "Scene" Then
            Key = Frame.Nodes.Item(N).Key
            FrameCoun = Frame.Nodes(Key).Children
            If FrameCoun <> 0 Or Frame.Nodes(Key).Tag = 4 Then
                nn = nn + 1
                Load frmMain.MnuScene(nn)
                frmMain.MnuScene(0).Visible = False
                frmMain.MnuScene(1).Checked = True
                frmMain.MnuScene(nn).Visible = True
                frmMain.MnuScene(nn).Enabled = True
                If Frame.Nodes(Key).Tag = 4 Then
                    frmMain.MnuScene(nn).Caption = Frame.Nodes(Key).Text
                Else
                    frmMain.MnuScene(nn).Caption = Frame.Nodes(Key).Text & "  (" & FrameCoun & ")"
                End If
                frmMain.MnuScene(nn).Tag = Frame.Nodes(Key).Key
            End If
        End If
        
    Next N
    If nn = 0 Then
        frmMain.MnuScene(0).Visible = True
        frmMain.MnuScene(0).Enabled = False
    End If
End Sub


Private Sub Frame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Frame.SelectedItem.Tag
   Case 1
    popup(1).Enabled = True
    popup(2).Enabled = False
    popup(3).Enabled = False
    popup(5).Caption = "Add 1 scene": popup(5).Enabled = True
    popup(6).Caption = "Add 5 scenes": popup(6).Enabled = True
    popup(7).Enabled = True
   Case 2
    popup(1).Enabled = True
    popup(2).Enabled = True
    popup(3).Enabled = True
    popup(5).Caption = "Add 1 frame": popup(5).Enabled = True
    popup(6).Caption = "Add 5 frames": popup(6).Enabled = True
    popup(7).Enabled = False
   Case 3
    popup(1).Enabled = True
    popup(2).Enabled = True
    popup(3).Enabled = False
    popup(5).Enabled = False
    popup(6).Enabled = False
    popup(7).Enabled = False
   Case 4
    popup(1).Enabled = True
    popup(2).Enabled = True
    popup(3).Enabled = False
    popup(5).Enabled = False
    popup(6).Enabled = False
    popup(7).Enabled = False
   End Select
 If Button = 2 Then
  PopupMenu Popop
 End If
End Sub


Private Sub popup_Click(Index As Integer)
    Select Case Index
        Case 1
            If Frame.SelectedItem.Tag <> 1 Then
                X = InputBox("Enter a new name for this scene", "Change name", Frame.SelectedItem.Text)
                If X <> "" Then
                    Frame.Nodes(Frame.SelectedItem.Key).Text = X
                    For N = 1 To cstTotalScenes
                        If Scenes(N).Key = Frame.SelectedItem.Key Then
                            Scenes(N).Name = X
                        End If
                    Next N
                End If
                UpdateToolbarlist
            Else
                X = InputBox("Enter a new name for the model", "Change name", Model.ProjectName)
                If X <> "" Then
                    frmMain.Joints.Nodes("Model").Text = X
                    sceneEdit.Frame.Nodes("Model").Text = X
                    frmMain.Caption = App.Title & " [" & Model.ProjectName & "]"
                    Model.ProjectName = X
                End If
            End If
         Case 2
            If Frame.SelectedItem.Tag = 2 Then
                Scn = Val(Mid(Frame.SelectedItem.Key, 6, 3))
                Scenes(Scn).Used = False
                Scenes(Scn).Name = ""
                Scenes(Scn).Key = ""
            End If
            Frame.Nodes.Remove Frame.SelectedItem.Key
        Case 3
            If Frame.SelectedItem.Tag = 2 Then
                Frame.Nodes.Remove Frame.SelectedItem.Key
            End If
        Case 5
            If Frame.SelectedItem.Tag = 1 Then Add_Scene "New Scene"
            If Frame.SelectedItem.Tag = 2 Then Add_Frame
        Case 6
            For N = 1 To 5
                If Frame.SelectedItem.Tag = 1 Then Add_Scene "New Scene"
                If Frame.SelectedItem.Tag = 2 Then Add_Frame
            Next N
        Case 7
            Add_Increment ("New Scene")
    End Select
End Sub

Private Sub Add_Increment(NewName)
    Frame.Nodes("Model").Expanded = True
    Add = FindScene
    If FindScene = 0 Then
        MsgBox "Sorry, you can only have " & cstTotalScenes & " frames": Exit Sub
    End If
    Key = "Scene" & Add
    Frame.Nodes.Add "Model", 4, Key, NewName, 4
    Frame.Nodes(Key).Tag = 4
    Scenes(Add).Key = Key
    Scenes(Add).Name = NewName
    Scenes(Add).Used = True
    Scenes(Add).Mode = 2
    UpdateToolbarlist
End Sub


Private Sub Timer1_Timer()
    If SBar.Tag = 0 Then Exit Sub
    If Frame.SelectedItem.Tag = 1 Then Exit Sub
    If Frame.SelectedItem.Tag = 2 Then
        If Frame.SelectedItem.Children = 0 Then Exit Sub
        GetThis = Frame.SelectedItem.Child.Key
        Frame.Nodes(GetThis).Selected = True
        Exit Sub
    End If
    If Frame.SelectedItem.Key = Frame.SelectedItem.LastSibling.Key Then
        Frame.Nodes(Frame.SelectedItem.FirstSibling.Key).Selected = True
        GetThis = Frame.SelectedItem.Key
        Label3.Caption = "Working file - " & GetThis
    Else
        GetThis = Frame.SelectedItem.Next.Key
        Frame.Nodes(GetThis).Selected = True
        Label3.Caption = "Working file - " & GetThis
    End If
End Sub

