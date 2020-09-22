VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 95/98"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   5
      Top             =   1800
      Width           =   2280
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Top             =   2280
      Width           =   2445
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Magic Mushroom Co"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "This product is freeware provided by"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -360
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AnimationShop 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   5280
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magic Mushroom Co"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   3585
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   120
      Picture         =   "Frmsplash.frx":0000
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor
    lblProductName.Caption = App.Title & " " & App.Major
End Sub



