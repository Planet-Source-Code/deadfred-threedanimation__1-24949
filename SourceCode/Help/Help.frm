VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Animation Shop 5 - Help"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   8940
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":0556
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text 
      Height          =   2175
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3836
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Help.frx":09AA
   End
   Begin MSComctlLib.TreeView Topics 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8281
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   471
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu popup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu pop 
         Caption         =   "New section"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldTopic As String
Dim Editable As Boolean
Const Path As String = "Help"

Private Sub Form_Load()
   ' Editable = True
    HelpTopicToLoad = ""
    On Error GoTo NoTopic
    Open App.Path & "\HelpTopic.txt" For Input As #1
        Input #1, HelpTopicToLoad
    Close
    Kill App.Path & "\HelpTopic.txt"
    If HelpTopicToLoad = "EditMeNow!!!" Then Editable = True: HelpTopicToLoad = "Main"
NoTopic:
    If HelpTopicToLoad = "" Then
        HelpTopicToLoad = "Main"
    End If
    UpdateTree
    On Error GoTo NotFound
    Topics.Nodes("Main").Selected = True
    Topics.Nodes(HelpTopicToLoad).Selected = True
    GetData
    If Editable = True Then
        Text.Locked = False
        Form2.Visible = True
    End If
    Exit Sub
NotFound:
    Ar = Ar & "Error - Help cannot find the files you want to look at. This may be because "
    Ar = Ar & "you have deleted, renamed or moved cirtain files. The help files should be contained "
    Ar = Ar & "in a folder named '" & Path & "' in the folder containing the Help exe. This folder should contain "
    Ar = Ar & "a number of files with the extenstion DAT, and a single TXT file named Tree.txt. If you have "
    Ar = Ar & "edited any of these files, it may also cause this program to experience errors"
    Text = Ar
    Topics.Enabled = False
End Sub

Private Sub Form_Resize()
    Topics.Height = Me.ScaleHeight
    Text.Height = Me.ScaleHeight
    Text.Width = Me.ScaleWidth - Topics.Width
    Text.Left = Topics.Width
End Sub

Private Sub UpdateTree()
    On Error GoTo NotWorking
    Topics.Nodes.Clear
    Topics.Nodes.Add , , "Main", "Animation Shop 5", 1
    Open App.Path & "\" & Path & "\tree.txt" For Input As #1
    Do
        Input #1, nrel
        Input #1, nName
        Input #1, nKey
        Topics.Nodes.Add nrel, 4, nKey, nName, 1
        Topics.Nodes(nrel).Image = 2
        If nrel = "Main" Then Topics.Nodes(nKey).EnsureVisible
    Loop While EOF(1) = False
    Exit Sub
NotWorking:
    Ar = Ar & "Error - Help read the file Tree.TXT. It may be because it has been altered, renamed, "
    Ar = Ar & "moved or deleted. Make sure that an unaltered copy of this file is present in the '" & Path & "' "
    Ar = Ar & "folder (The folder with all the DAT) files in..."
    Text = Ar
Topics.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form2
    Unload Me
    End
End Sub

Private Sub pop_Click(Index As Integer)
Select Case Index
    Case 1
        Close
        Open App.Path & "\" & Path & "\tree.txt" For Append As #1
            x = InputBox("Enter a topic name")
            If x = "" Then Exit Sub
            Print #1, """"; Topics.SelectedItem.Key; """"; ",";
            Print #1, """"; x; """"; ",";
            Print #1, """"; x; """"
            Topics.Nodes.Add Topics.SelectedItem.Key, 4, x, x, 1
            Topics.Nodes(x).EnsureVisible
        Close
End Select
End Sub

Private Sub Topics_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Editable = False Then Exit Sub
    If Button = 2 Then PopupMenu popup
End Sub

Private Sub Topics_NodeClick(ByVal Node As MSComctlLib.Node)
    GetData
End Sub

Private Sub GetData()
    On Error Resume Next
    If OldTopic <> "" Then Text.SaveFile App.Path & "\" & Path & "\" & OldTopic & ".dat"
    OldTopic = Topics.SelectedItem.Key
    Text.TextRTF = ""
    Text.LoadFile App.Path & "\" & Path & "\" & OldTopic & ".dat"
End Sub
