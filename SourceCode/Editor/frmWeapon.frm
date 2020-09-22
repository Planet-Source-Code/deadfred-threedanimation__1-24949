VERSION 5.00
Begin VB.Form frmWeapon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Weapon Types"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "D&elete"
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtModelID 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ListBox txtContact 
         Height          =   255
         ItemData        =   "frmWeapon.frx":0000
         Left            =   3000
         List            =   "frmWeapon.frx":000D
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Help"
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ListBox txtTrail 
         Height          =   255
         ItemData        =   "frmWeapon.frx":002A
         Left            =   3000
         List            =   "frmWeapon.frx":0040
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Model ID"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "&On contact"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "&Damage"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "&Trail"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "&ID"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "&Speed"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "&Name"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.ListBox lstAll 
      Height          =   1230
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "&Existing weapon types"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WeaponON As Byte

Private Sub cmdSave_Click()
    If IsThisValid(txtID, 0) = False Then
        MsgBox "Incorrect ID value", 16, "Data Error"
        Exit Sub
    End If
    If IsThisValid(txtSpeed, 1) = False Then
        MsgBox "Incorrect Speed value", 16, "Data Error"
        Exit Sub
    End If
    If IsThisValid(txtDamage, 0) = False And Val(txtDamage) <> 0 Then
        MsgBox "Incorrect Damage value", 16, "Data Error"
        Exit Sub
    End If
    If IsThisValid(txtModelID, 0) = False Then
        MsgBox "Incorrect Model ID value", 16, "Data Error"
        Exit Sub
    End If
    Close
    Open frmSettings.txtWeaponFile For Input As #1
    Open frmSettings.txtWeaponFile & "temp" For Output As #2
    If EOF(1) = True Then
    Else
        OnNow = 0
        Do
            OnNow = OnNow + 1
            Input #1, WeaponName
            Input #1, Speed
            Input #1, ID
            Input #1, trail
            Input #1, Damage
            Input #1, Contact
            Input #1, ModelID
            If OnNow = WeaponON Then
                Print #2, txtName
                Print #2, txtSpeed
                Print #2, txtID
                Print #2, txtTrail.List(txtTrail.ListIndex)
                Print #2, txtDamage
                Print #2, txtContact.List(txtContact.ListIndex)
                Print #2, txtModelID
            Else
                Print #2, WeaponName,
                Print #2, Speed
                Print #2, ID
                Print #2, trail
                Print #2, Damage
                Print #2, Contact
                Print #2, ModelID
            End If
        Loop While EOF(1) <> True
    End If
    If OnNow + 1 = WeaponON Then
        Print #2, txtName
        Print #2, txtSpeed
        Print #2, txtID
        Print #2, txtTrail.Text
        Print #2, txtDamage
        Print #2, txtContact.Text
        Print #2, txtModelID
    End If
    Close
    Kill frmSettings.txtWeaponFile
    Name frmSettings.txtWeaponFile & "temp" As frmSettings.txtWeaponFile
    Open frmSettings.txtWeaponFile For Input As #1
    lstAll.Clear
    If EOF(1) = True Then
    Else
        Do
            Input #1, WeaponName
            Input #1, Speed
            Input #1, ID
            Input #1, trail
            Input #1, Damage
            Input #1, Contact
            Input #1, ModelID
            lstAll.AddItem WeaponName
        Loop While EOF(1) <> True
    End If
    Close
    lstAll.AddItem "[New]"
    lstAll.Selected(WeaponON - 1) = True
End Sub

Private Sub Command1_Click()
        If ShowHelp("New weapons") = False Then
            MsgBox "help files are missing, or fucked!"
        End If
End Sub

Private Sub Command2_Click()
    If lstAll.Text = "[New]" Then Exit Sub
    Open frmSettings.txtWeaponFile For Input As #1
    Open frmSettings.txtWeaponFile & "temp" For Output As #2
    OnNow = 0
    Do
        OnNow = OnNow + 1
        Input #1, WeaponName
        Input #1, Speed
        Input #1, ID
        Input #1, trail
        Input #1, Damage
        Input #1, Contact
        Input #1, ModelID
        If OnNow = WeaponON Then
        Else
            Print #2, WeaponName
            Print #2, Speed
            Print #2, ID
            Print #2, trail
            Print #2, Damage
            Print #2, Contact
            Print #2, ModelID
        End If
    Loop While EOF(1) <> True
    Close
    Kill frmSettings.txtWeaponFile
    Name frmSettings.txtWeaponFile & "temp" As frmSettings.txtWeaponFile
    Close
    Open frmSettings.txtWeaponFile For Input As #1
    lstAll.Clear
    If EOF(1) = True Then
    Else
        Do
            Input #1, WeaponName
            Input #1, Speed
            Input #1, ID
            Input #1, trail
            Input #1, Damage
            Input #1, Contact
            Input #1, ModelID
            lstAll.AddItem WeaponName
        Loop While EOF(1) <> True
    End If
    Close
    lstAll.AddItem "[New]"
    lstAll.Selected(WeaponON - 1) = True
End Sub

Private Sub Form_Load()
    Close
    Open frmSettings.txtWeaponFile For Input As #1
    lstAll.Clear
    If EOF(1) = True Then
    Else
        Do
            Input #1, WeaponName
            Input #1, Speed
            Input #1, ID
            Input #1, trail
            Input #1, Damage
            Input #1, Contact
            Input #1, ModelID
            lstAll.AddItem WeaponName
        Loop While EOF(1) <> True
    End If
    Close
    lstAll.AddItem "[New]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmJointProp.Enabled = True
    frmJointProp.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub lstAll_Click()
    WeaponON = lstAll.ListIndex + 1
    Close
    If lstAll.ListIndex + 1 = lstAll.ListCount Then
        txtName = "New"
        txtSpeed = ""
        txtID = ""
        txtTrail = ""
        txtDamage = ""
        txtContact = ""
        txtModelID = ""
        txtTrail.Selected(0) = True
        txtContact.Selected(0) = True
    Else
    Open frmSettings.txtWeaponFile For Input As #1
        OnNow = 0
        Do
            OnNow = OnNow + 1
            Input #1, WeaponName
            Input #1, Speed
            Input #1, ID
            Input #1, trail
            Input #1, Damage
            Input #1, Contact
            Input #1, ModelID
            If OnNow = WeaponON Then
                txtName = WeaponName
                txtSpeed = Speed
                txtID = ID
                txtTrail.Text = trail
                txtDamage = Damage
                txtContact.Text = Contact
                txtModelID = ModelID
            End If
        Loop While EOF(1) <> True
    End If
    Close
End Sub
