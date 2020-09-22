VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Model..."
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "FrmExport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Frame Frame 
      Height          =   4695
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox ChkBotDetail 
         Caption         =   "Include Bot Details"
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chKweapon 
         Caption         =   "Include weapons"
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.OptionButton Mode 
         Caption         =   "Object with skeliton - Faces on seperate lines"
         Height          =   375
         Index           =   6
         Left            =   1800
         TabIndex        =   20
         Top             =   3120
         Value           =   -1  'True
         Width           =   3855
      End
      Begin VB.OptionButton Mode 
         Caption         =   "Single Object - Faces on seperate lines"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   19
         Top             =   2400
         Width           =   3135
      End
      Begin VB.OptionButton Mode 
         Caption         =   "Seperate objects - Faces on seperate lines"
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   18
         Top             =   1560
         Width           =   3495
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         ItemData        =   "FrmExport.frx":030A
         Left            =   6600
         List            =   "FrmExport.frx":0311
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chkNotes 
         Caption         =   "Include notes, staring with '//'"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   4200
         Width           =   2535
      End
      Begin VB.OptionButton Mode 
         Caption         =   "Object with skeliton - All faces on one line"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   4
         Top             =   2880
         Width           =   3375
      End
      Begin VB.CheckBox chkStartAt 
         Caption         =   "Vertext list starts at '1'"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   4200
         Width           =   1935
      End
      Begin VB.OptionButton Mode 
         Alignment       =   1  'Right Justify
         Caption         =   "No compression"
         Height          =   255
         Index           =   0
         Left            =   320
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Mode 
         Caption         =   "Seperate objects - All faces on one line"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   3255
      End
      Begin VB.OptionButton Mode 
         Caption         =   "Single Object - All faces on one line"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   360
         X2              =   6360
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   360
         X2              =   6360
         Y1              =   3735
         Y2              =   3735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   240
         X2              =   6240
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   240
         X2              =   6240
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Compressed file"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      Height          =   4695
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.OptionButton StoreMode 
         Caption         =   "Export to one file"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   11
         Top             =   3240
         Width           =   1935
      End
      Begin VB.OptionButton StoreMode 
         Caption         =   "Export to seperate files"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   10
         Top             =   1440
         Value           =   -1  'True
         Width           =   2055
      End
      Begin MSComctlLib.TreeView Frames 
         Height          =   3735
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   6588
         _Version        =   393217
         Style           =   6
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Select the frames you want to export"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2775
      End
   End
   Begin MSComctlLib.TabStrip Tabs 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9551
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Format"
            Key             =   "Key1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Frames"
            Key             =   "Key2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub cmdCancel_Click()
    Timei.Visible = False
    Unload Me
End Sub

Private Sub cmdCompile_Click()
    
    FileName = SetFileName("Compile", "Export file to..")
    
    If FileName <> "" Then
        If CheckOverwrite(FileName) = False Then Exit Sub
        CompilerControler FileName, 1
    End If
    
    
 End Sub
    
Public Sub CompilerControler(FileName, BatchOn)
    
    SetUPSkeliton
    
    Timei.Visible = True
    Timei.TimeLeft.ToolTipText = "Click this bar to abort export process"
    Timei.Caption = "Exporting model - Please wait"
    Me.Visible = False
    
    
    lent = Len(FileName)
    Extention = Mid(FileName, lent - 2, 3)
    FileName = Mid(FileName, 1, lent - 4)
    
    
    Destroy FileName & n & "." & Extention
    
    ticked = 0
    For n = 1 To Frames.Nodes.Count
        If Frames.Nodes.Item(n).Tag = 1 Then
            If Frames.Nodes.Item(n).Checked = True Then
                ticked = ticked + 1
            End If
        End If
    Next n
    NoFrames = False
    If ticked = 0 Then
        ticked = 1
        NoFrames = True
        StoreMode(1).Value = True
    End If
    
    For n = 1 To ticked
        If NoFrames = False Then
            Do: Fin = Fin + 1: Loop While Frames.Nodes.Item(Fin).Tag <> 1 Or Frames.Nodes.Item(Fin).Checked = False
        End If
        If StoreMode(0).Value = True Then Open FileName & n & "." & Extention For Output As #1
        If StoreMode(1).Value = True Then Open FileName & "." & Extention For Append As #1
        If NoFrames = False Then
            Print #1, Frames.Nodes.Item(Fin).Key; ",     ";
            lenf = Len(Frames.Nodes.Item(Fin).Key)
            sceneon = Val(Mid(Frames.Nodes.Item(Fin).Key, 6, lenf - 7))
            Print #1, sceneEdit.Frame.Nodes("Scene" & sceneon).Text
        End If
        
        
        If ChkBotDetail = 1 Then
            Print #1, "ID, " & frmProperties.txtBotID
            Print #1, "Discription, """ & Model.BotDis; """"
            Print #1, "Cost, " & Model.BotCost
            Print #1, "Weight, " & Model.BotWeight
        End If
        
        
        If Mode(0).Value = True Then Compile.SingleFaceExport
        
        If Mode(2).Value = True Then Compile.SingleObjectExport
        If Mode(3).Value = True Then Compile.SingleObjectMultiLineExport
        
        If Mode(1).Value = True Then Compile.CompleteObject
        If Mode(5).Value = True Then Compile.CompleteObject
        
        If Mode(4).Value = True Then Compile.CompleteObject
        If Mode(6).Value = True Then Compile.CompleteObject
        
        
        Close #1
    Next n
    
    Timei.Visible = False
    frmMain.Enabled = True
    If frmMain.Visible = True Then frmMain.SetFocus
    
    If BatchOn = 1 Then
        Ms = Ms & "Your model has been successfully complied to" & vbNewLine & vbNewLine
        If StoreMode(0).Value = True Then
            Ms = Ms & FileName & 1 & "." & Extention & vbNewLine & vbNewLine
            Ms = Ms & "to" & vbNewLine & vbNewLine
            Ms = Ms & FileName & ticked & "." & Extention
        End If
        If StoreMode(1).Value = True Then Ms = Ms & FileName & "." & Extention
        Ms = Ms & vbNewLine & vbNewLine & "As this file is all text, you can view it in note pad, or another text editor. Would you like to open it now?"
        Responce = MsgBox(Ms, 64 + vbYesNo)
        If Responce = 6 Then
            X = "notepad.exe"
            c = FileName & "." & Extention
            Shell "notepad.exe " + c, vbNormalFocus
        End If
    End If
    Exit Sub
    
End Sub

Private Sub cmHelp_Click()
    If Frame(0).Visible = True Then
         ShowHelp "Exporting models"
    Else
        ShowHelp "Exporting frames"
    End If
End Sub

Private Sub Form_Load()
    
    
        XX = Mid(Model.ProjectFileName, 1, Len(Model.ProjectFileName) - 4) & ".dat"
        frmMain.GetFile.FileName = XX
    
    
    Frame(0).Visible = True
    Frames.Nodes.Add , 0, "Model", "Complete model"
    For n = 1 To cstTotalScenes
        If Scenes(n).Used = True Then
            Frames.Nodes.Add "Model", 4, Scenes(n).Key, Scenes(n).Name
            Frames.Nodes(Scenes(n).Key).Tag = 2
            For m = 1 To sceneEdit.Frame.Nodes(Scenes(n).Key).Children
                Frames.Nodes.Add Scenes(n).Key, 4, Scenes(n).Key & "_" & m, "Frame " & m
                Frames.Nodes(Scenes(n).Key & "_" & m).EnsureVisible
                Frames.Nodes(Scenes(n).Key & "_" & m).Tag = 1
            Next m
            FrameOn = ""
        End If
    Next n
    Frames.Nodes("Model").Checked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub Frames_NodeCheck(ByVal Node As MSComctlLib.Node)
    If Node.Key = "Model" Then
        For n = 1 To Frames.Nodes.Count
            Frames.Nodes.Item(n).Checked = Node.Checked
        Next n
    End If
    If Node.Tag = 2 Then
        For n = 1 To Node.Children
            Frames.Nodes(Node.Key & "_" & n).Checked = Node.Checked
        Next n
    End If
End Sub



Private Sub Mode_DblClick(Index As Integer)
    cmdCompile_Click
End Sub

Private Sub Tabs_Click()
    Frame(0).Visible = False
    Frame(1).Visible = False
    X = Mid(Tabs.SelectedItem.Key, 4, 1)
    Frame(X - 1).Visible = True
End Sub




