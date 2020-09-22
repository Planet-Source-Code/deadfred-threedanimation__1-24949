VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Timei 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporting model... Please wait"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar TimeLeft 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Click this bar to stop compiling"
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Percentage Done"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Timei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TimeLeft_Click()
    If TimeLeft.ToolTipText = "" Then Exit Sub
    X = MsgBox("Are you sure you want to stop exporting your model?", 292)
    If X = 6 Then
        If frmMain.Visible = True Then
            Timei.Visible = False
            frmMain.Enabled = True
            frmMain.Visible = True
            frmMain.SetFocus
        Else
            Timei.Visible = False
            Unload frmSplash
            Destroy "Shape.txt"
            End
        End If
    End If
End Sub
