VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Properties"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "Properties.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   4
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chKFrc 
         Caption         =   "Force visible skeliton"
         Height          =   195
         Left            =   1680
         TabIndex        =   38
         ToolTipText     =   "Display skeliton as part of the actual model"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtBotDiscription 
         Height          =   1005
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "Discription of item for the user"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtBotWeight 
         Height          =   285
         Left            =   1680
         TabIndex        =   35
         Text            =   "0"
         ToolTipText     =   "Weight of item"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtBotCost 
         Height          =   285
         Left            =   1680
         TabIndex        =   34
         Text            =   "0"
         ToolTipText     =   "Price of robot peice"
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtBotID 
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Text            =   "18394556575354"
         ToolTipText     =   "Verification of file"
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Discription"
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight"
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "ID"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   3
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkShowatStart 
         Caption         =   "Show notes whenever the model is loaded"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         ToolTipText     =   "If ticked, the following text appears when the file is opened"
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   720
         Width           =   5055
      End
   End
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkSideBar 
         Caption         =   "Display Side Bar"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ListBox lstScene 
         Height          =   1425
         Left            =   2880
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkRunAnim 
         Caption         =   "Run Animation"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ListBox lstViewMode 
         Height          =   1425
         ItemData        =   "Properties.frx":0442
         Left            =   840
         List            =   "Properties.frx":0455
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "View mode:"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5775
      Begin VB.Label lblCreationDate 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   19
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Creation date"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblFrames 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   17
         Top             =   3600
         Width           =   750
      End
      Begin VB.Label lblScenes 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   3240
         Width           =   750
      End
      Begin VB.Label lblJoints 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   15
         Top             =   2880
         Width           =   750
      End
      Begin VB.Label lblverecies 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   14
         Top             =   2280
         Width           =   750
      End
      Begin VB.Label lblTotalFaces 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   13
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label lblObjects 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   12
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label lblmodelfilename 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   750
      End
      Begin VB.Label lblModelName 
         AutoSize        =   -1  'True
         Caption         =   "Unavalible"
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Total vertecies"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Faces in model"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Frames in model"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Scenes in model"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Joints in model"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Objects in model"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Model File Name"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Model Name"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Model Properties"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "Display statistics about your model"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Startup"
            Object.Tag             =   "2"
            Object.ToolTipText     =   "Define how your model is displayed when you open it..."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notes"
            Object.Tag             =   "3"
            Object.ToolTipText     =   "Store text discibing the model"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "FastTank Properties"
            Object.Tag             =   "4"
            Object.ToolTipText     =   "Properties for the game FastTank"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub chkRunAnim_Click()
    If Me.chkRunAnim.Value = 0 Then Model.StartAnimated = False
    If Me.chkRunAnim.Value = 1 Then Model.StartAnimated = True
End Sub

Private Sub chkShowatStart_Click()
    If Me.chkShowatStart.Value = 0 Then Model.ShowNotesAtStart = False
    If Me.chkShowatStart.Value = 1 Then Model.ShowNotesAtStart = True
End Sub

Private Sub chkSideBar_Click()
    If Me.chkSideBar.Value = 0 Then Model.StartShowSideBar = False
    If Me.chkSideBar.Value = 1 Then Model.StartShowSideBar = True
End Sub


Private Sub Form_Load()
    Me.lblModelName = Model.ProjectName
    Me.lblmodelfilename = Model.ProjectFileName
    Me.lblObjects = CountObjects
    Me.lblTotalFaces = CountFaces
    Me.lblverecies = CountVertecies
    Me.lblJoints = CountJoints
    Me.lblScenes = CountScenes
    Me.lblCreationDate = Model.ProjectCreated
    
    Me.txtBotID = Model.BotID
    Me.txtBotCost = Model.BotCost
    Me.txtBotDiscription = Model.BotDis
    Me.txtBotWeight = Model.BotWeight
    
    If CountScenes <> 0 Then
        For n = 1 To cstTotalScenes
            If Scenes(n).Used = True Then
                Me.lstScene.AddItem Scenes(n).Name
            End If
        Next n
        Me.lstScene.Selected(Model.StartSceneName) = True
    End If
    
    If Model.ShowNotesAtStart = False Then Me.chkShowatStart.Value = 0
    If Model.ShowNotesAtStart = True Then Me.chkShowatStart.Value = 1
    
    txtNotes = Model.ModelNotes
    
    If Model.StartAnimated = False Then Me.chkRunAnim.Value = 0
    If Model.StartAnimated = True Then Me.chkRunAnim.Value = 1
    
    If Model.BotForce = False Then Me.chKFrc.Value = 0
    If Model.BotForce = True Then Me.chKFrc.Value = 1
   
    
    Me.lstViewMode.Selected(Model.StartViewMode) = True
    If Model.StartShowSideBar = False Then Me.chkSideBar.Value = 0
    If Model.StartShowSideBar = True Then Me.chkSideBar.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    
    Model.BotCost = Me.txtBotCost
    Model.BotDis = Me.txtBotDiscription
    Model.BotWeight = Me.txtBotWeight
    Model.BotID = Me.txtBotID
    Model.BotForce = Me.chKFrc

End Sub

Private Sub lstScene_Click()
    Model.StartSceneName = lstScene.ListIndex
End Sub

Private Sub lstViewMode_Click()
    Model.StartViewMode = lstViewMode.ListIndex
End Sub

Private Sub TabStrip1_Click()
    For n = 1 To Frame.Count: Frame(n).Visible = False: Next n
    Frame(TabStrip1.SelectedItem.Tag).Visible = True
End Sub


Private Sub txtNotes_Change()
    Model.ModelNotes = txtNotes
End Sub
