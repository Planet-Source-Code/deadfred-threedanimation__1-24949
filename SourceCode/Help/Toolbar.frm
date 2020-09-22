VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ToolBar"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3255
   Begin VB.ComboBox TextSize 
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   630
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1111
      _Version        =   393216
      LargeChange     =   1
      Max             =   2
      TickStyle       =   2
   End
   Begin VB.ListBox Fonts 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click(Index As Integer)
If Check1(Index).Value = 1 Then
    If Index = 0 Then Form1.Text.SelBold = True
    If Index = 1 Then Form1.Text.SelUnderline = True
    If Index = 2 Then Form1.Text.SelItalic = True
End If
If Check1(Index).Value = 0 Then
    If Index = 0 Then Form1.Text.SelBold = False
    If Index = 1 Then Form1.Text.SelUnderline = False
    If Index = 2 Then Form1.Text.SelItalic = False
End If
End Sub

Private Sub Form_Load()
    For n = 1 To Screen.FontCount - 1
        Fonts.AddItem Screen.Fonts(n)
    Next n
    TextSize.AddItem "8"
    TextSize.AddItem "10"
    TextSize.AddItem "11"
    TextSize.AddItem "12"
    TextSize.AddItem "14"
    TextSize.AddItem "16"
    TextSize.AddItem "18"
    TextSize.AddItem "20"
    TextSize.AddItem "22"
    TextSize.AddItem "24"
    TextSize.AddItem "28"
    TextSize.AddItem "32"
End Sub

Private Sub List1_Click()
Form1.Text.SelFontName = List1.Text
End Sub

Private Sub Slider1_Click()
If Slider1.Value = 0 Then Form1.Text.SelAlignment = 0
If Slider1.Value = 1 Then Form1.Text.SelAlignment = 2
If Slider1.Value = 2 Then Form1.Text.SelAlignment = 1
End Sub

Private Sub TextSize_Change()
Form1.Text.SelFontSize = TextSize
End Sub

Private Sub TextSize_Click()
Form1.Text.SelFontSize = TextSize
End Sub
