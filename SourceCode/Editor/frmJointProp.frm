VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJointProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Joint properties"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmJointProp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   4455
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      Begin VB.TextBox MoreText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox MainData 
         BackColor       =   &H00FFFFFF&
         Height          =   3615
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Linked to"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame 
      Height          =   4455
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ComboBox cmbWeaponType 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Text            =   "Cannon"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.ComboBox cmbWeaponName 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Text            =   "New Weapon"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CheckBox chkWeapon 
         Caption         =   "This joint is a weapon"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Weapon Type"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   3840
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   3840
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Weapon name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
   End
   Begin MSComctlLib.TabStrip Tabz 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Joint properties"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "Details about the joint"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Weapons"
            Object.Tag             =   "2"
            Object.ToolTipText     =   "Sets the joint as a weapon"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmJointProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdHelp_Click()
        If ShowHelp("Joint settings") = False Then
            MsgBox "help files are missing, or fucked!"
        End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub



Private Sub Tabz_Click()
    For n = 1 To Frame.Count
        Frame(n - 1).Visible = False
        Frame(Tabz.SelectedItem.Tag - 1).Visible = True
    Next n
End Sub

Private Sub Form_Unload(Cancel As Integer)
    For n = 1 To cstTotalJoints
        If BaseFrame(n).Selected = True Then
            If chkWeapon = 1 Then
                BaseFrame(n).IsAWeapon = True
                BaseFrame(n).WeaponName = cmbWeaponName
                BaseFrame(n).WeaponType = cmbWeaponType
            End If
        End If
    Next n
    frmMain.Enabled = True
    frmMain.SetFocus
    frmMain.DrawModel
End Sub

Public Sub RunAtStart()

    MoreText.Top = MainData.Top + 30
    MoreText.Height = MainData.Height - 60
    
    For n = 1 To cstTotalJoints
        If BaseFrame(n).Used = True And BaseFrame(n).IsAWeapon = True Then
            cmbWeaponName.AddItem BaseFrame(n).WeaponName
            cmbWeaponType.AddItem BaseFrame(n).WeaponType
        End If
    Next n
    If CountSelectedJoints <> 1 Then
        Caption = "Joint properties [ Multiple joints ]"
        For n = 1 To cstTotalJoints
            If BaseFrame(n).Selected = True Then
                MainData = MainData & BaseFrame(n).Name & vbNewLine
                MoreText = MoreText & BaseFrame(FindTarget(BaseFrame(n).Target)).Name & vbNewLine
                jType = BaseFrame(n).JointType
                If BaseFrame(n).IsAWeapon = True Then
                    chkWeapon.Value = 1
                    cmbWeaponName = ""
                    LstWeaponType = ""
                End If
            End If
        Next n
    Else
        For n = 1 To cstTotalJoints
            If BaseFrame(n).Selected = True Then
                Caption = "Joint properties [ " & BaseFrame(n).Name & " ]"
                MainData = MainData & BaseFrame(n).Name & vbNewLine
                MoreText = MoreText & BaseFrame(FindTarget(BaseFrame(n).Target)).Name & vbNewLine
                If BaseFrame(n).IsAWeapon = True Then
                    chkWeapon.Value = 1
                    cmbWeaponName = BaseFrame(n).WeaponName
                    cmbWeaponType = BaseFrame(n).WeaponType
                End If
            End If
        Next n
    End If
End Sub

