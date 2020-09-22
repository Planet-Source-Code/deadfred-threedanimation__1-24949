VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "Animation Shop 6 - []"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   -75
   ClientWidth     =   10890
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox DirectX 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   60
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox View 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
      Begin VB.PictureBox Corner 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   67
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton Block 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   2640
         Width           =   255
      End
      Begin VB.VScrollBar UDBar 
         Height          =   2655
         LargeChange     =   100
         Left            =   4080
         Max             =   1000
         Min             =   -1000
         SmallChange     =   100
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.HScrollBar LRBar 
         Height          =   255
         LargeChange     =   100
         Left            =   0
         Max             =   1000
         Min             =   -1000
         SmallChange     =   100
         TabIndex        =   9
         Top             =   2640
         Width           =   4095
      End
      Begin VB.PictureBox SideRule 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   0
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   68
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox TopRule 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   257
         TabIndex        =   71
         Top             =   0
         Width           =   3855
      End
      Begin VB.PictureBox Texture 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   240
         ScaleHeight     =   143
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   191
         TabIndex        =   92
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Line Guide 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         DrawMode        =   2  'Blackness
         Tag             =   "False"
         Visible         =   0   'False
         X1              =   160
         X2              =   224
         Y1              =   104
         Y2              =   104
      End
      Begin VB.Line Axis 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         DrawMode        =   4  'Mask Not Pen
         Tag             =   "X"
         Visible         =   0   'False
         X1              =   184
         X2              =   184
         Y1              =   0
         Y2              =   56
      End
   End
   Begin MSComctlLib.TabStrip MainTab 
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6800
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Top View"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "View your model from the top"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Front View"
            Object.Tag             =   "2"
            Object.ToolTipText     =   "View your model from the side"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Side View"
            Object.Tag             =   "3"
            Object.ToolTipText     =   "View your model from the front"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "3D Wireframe"
            Object.Tag             =   "4"
            Object.ToolTipText     =   "View your model in 3D"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Model Preview"
            Object.Tag             =   "5"
            Object.ToolTipText     =   "Preview your model"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Texture Map"
            Object.Tag             =   "6"
            Object.ToolTipText     =   "Display the models texture vertecies and face"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmaLightEffect 
      Interval        =   10
      Left            =   480
      Tag             =   "0"
      Top             =   5640
   End
   Begin VB.TextBox CurrentSideBar 
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ShowTime 
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   8040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Sbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6705
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "14:42"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13573
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CBar 
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   794
      BandBorders     =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   10095
      _CBHeight       =   450
      _Version        =   "6.0.8169"
      Child1          =   "TBar1"
      MinHeight1      =   390
      Width1          =   3255
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Child2          =   "TBar3"
      MinHeight2      =   390
      Width2          =   3120
      NewRow2         =   0   'False
      Child3          =   "TBar2"
      MinHeight3      =   375
      Width3          =   3495
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar TBar3 
         Height          =   390
         Left            =   3420
         TabIndex        =   7
         Top             =   30
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "MisqStuff"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Zoom out"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Zoom in"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Stop animation"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Play animation"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Enlarge View window"
               ImageIndex      =   3
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TBar2 
         Height          =   375
         Left            =   6540
         TabIndex        =   6
         Top             =   30
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   661
         ButtonWidth     =   609
         ButtonHeight    =   556
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "MyIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Select"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Create new brush"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Edit brush"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Scale Brush"
               ImageIndex      =   6
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Rotate Brush"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Add joint"
               ImageIndex      =   4
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TBar1 
         Height          =   390
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "StandadIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Start a new model"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open an existing model"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Save the current model"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy the selected brushes to memory"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Paste the contents of the clip board"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Moves the selected brushes to memory"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList StandadIcons 
      Left            =   1080
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0532
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0646
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":075A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":086E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList MyIcons 
      Left            =   2640
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0982
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1202
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":183A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList MisqStuff 
      Left            =   1680
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3016
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":346A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3786
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":434A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog GetFile 
      Left            =   480
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      HelpKey         =   "k,kkkk"
      MaxFileSize     =   240
   End
   Begin VB.CommandButton SideBar 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "The Side Bar"
      Top             =   5640
      Width           =   585
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame"
      Height          =   5535
      Index           =   6
      Left            =   5640
      TabIndex        =   53
      Top             =   600
      Width           =   2895
      Begin VB.OptionButton optJoint 
         Caption         =   "Joint Objects"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   119
         ToolTipText     =   "Seperate a single object into two seperate objects"
         Top             =   4440
         Width           =   1815
      End
      Begin VB.OptionButton optSeperate 
         Caption         =   "Seperate faces"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   118
         ToolTipText     =   "Seperate a single object into two seperate objects"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.OptionButton optDelVertecies 
         Caption         =   "Delete vertecies"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   117
         ToolTipText     =   "Delete unwanted or unused vertecies from an object"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.OptionButton edMode 
         Caption         =   "Select"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   91
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton edMode 
         Caption         =   "Extend Face"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   90
         ToolTipText     =   "Extends the face away from the object"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.OptionButton edMode 
         Caption         =   "Bend Face"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   89
         ToolTipText     =   "Allows you to bend a face by draging its center"
         Top             =   3120
         Width           =   2175
      End
      Begin VB.OptionButton edMode 
         Caption         =   "Delete Face"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   88
         ToolTipText     =   "Removes the face from the object"
         Top             =   2760
         Width           =   2175
      End
      Begin VB.OptionButton optConvert 
         Caption         =   "Convert faces into Triangles"
         Height          =   255
         Left            =   240
         TabIndex        =   86
         ToolTipText     =   "Turns all faces the object into a serise of triangles"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.OptionButton edMode 
         Caption         =   "Move each face"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   56
         ToolTipText     =   "Moves the position of each face"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Flip Brush"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   73
         ToolTipText     =   "Sets the changes to your model"
         Top             =   4920
         Width           =   2055
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "Flip Brush horizontally"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   57
         ToolTipText     =   "Swaps the left and right sides of the object over"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "Flip Brush vertically"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   58
         ToolTipText     =   "Swaps the top and bottom of the object over"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optInvert 
         Caption         =   "Reverse Brush faces"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         ToolTipText     =   "Changes the direct each face is facing"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton edMode 
         Caption         =   "Move each vertex"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   55
         ToolTipText     =   "Moves te position of each vertex"
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton edMode 
         Caption         =   "Squew selected object"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame1"
      Height          =   6255
      Index           =   3
      Left            =   5640
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.OptionButton opSelectJ 
         Caption         =   "Select tool"
         Height          =   255
         Left            =   360
         TabIndex        =   114
         ToolTipText     =   "As thought you have the 'Select Tool' on the toolbar selected"
         Top             =   4320
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton opAddJ 
         Caption         =   "Add joint to model"
         Height          =   255
         Left            =   360
         TabIndex        =   113
         ToolTipText     =   "Creates a new joint whee you click"
         Top             =   4680
         Width           =   1695
      End
      Begin VB.OptionButton opChangeJ 
         Caption         =   "Change Target of joint"
         Height          =   255
         Left            =   360
         TabIndex        =   112
         ToolTipText     =   "Drag from one joint to another to set the target"
         Top             =   5040
         Width           =   2175
      End
      Begin MSComctlLib.TreeView Joints 
         Height          =   3975
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   7011
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "MyIcons"
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame1"
      Height          =   6255
      Index           =   2
      Left            =   5880
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox PresetScale 
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Text            =   "Text1"
         ToolTipText     =   "The size that the object will become compared to before the opperation"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.ListBox Scales 
         Height          =   1230
         ItemData        =   "frmMain.frx":479E
         Left            =   360
         List            =   "frmMain.frx":47C6
         TabIndex        =   38
         Top             =   3240
         Width           =   2295
      End
      Begin VB.OptionButton SklMode 
         Caption         =   "Custom scale"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton cmdScale 
         Caption         =   "Scale brush"
         Height          =   375
         Left            =   360
         TabIndex        =   30
         ToolTipText     =   "Click to set the changes to your model"
         Top             =   5520
         Width           =   2055
      End
      Begin MSComctlLib.Slider sclXDim 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   29
         ToolTipText     =   "The size of the object along the X axis"
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Min             =   50
         Max             =   150
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin MSComctlLib.Slider sclXDim 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   32
         ToolTipText     =   "The size of the object along the Y axis"
         Top             =   1440
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Min             =   50
         Max             =   150
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin MSComctlLib.Slider sclXDim 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   34
         ToolTipText     =   "The size of the object along the Z axis"
         Top             =   2040
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Min             =   50
         Max             =   150
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin VB.OptionButton SklMode 
         Caption         =   "Preset scale"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Z"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Y"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame2"
      Height          =   6255
      Index           =   5
      Left            =   8040
      TabIndex        =   48
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox chkVert 
         Caption         =   "Select vertecies"
         Height          =   255
         Left            =   360
         TabIndex        =   87
         ToolTipText     =   "Allows you to select individual vertecies within selected objects"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CommandButton cmdMove 
         Height          =   495
         Index           =   0
         Left            =   1080
         Picture         =   "frmMain.frx":47FD
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Move selected object up one unit"
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Height          =   495
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":4B07
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Move selected object left one unit"
         Top             =   4320
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Height          =   495
         Index           =   2
         Left            =   1080
         Picture         =   "frmMain.frx":4E11
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Move selected object down one unit"
         Top             =   4800
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Height          =   495
         Index           =   3
         Left            =   1560
         Picture         =   "frmMain.frx":511B
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Move selected object right one unit"
         Top             =   4320
         Width           =   495
      End
      Begin VB.CheckBox chkSelDesel 
         Caption         =   "Don't Select/Deselect"
         Height          =   375
         Left            =   360
         TabIndex        =   62
         ToolTipText     =   "Prevents you from selecting or deselecting objects"
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CheckBox chkTotalSelect 
         Caption         =   "Box band must be complete"
         Height          =   375
         Left            =   360
         TabIndex        =   61
         ToolTipText     =   "Objects must be completely surounded to be selected"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CheckBox chkAttchOb 
         Caption         =   "Select all attached objects"
         Height          =   255
         Left            =   360
         TabIndex        =   52
         ToolTipText     =   "Selects objects attached to selected joints"
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkJoints 
         Caption         =   "Select joints"
         Height          =   375
         Left            =   360
         TabIndex        =   51
         ToolTipText     =   "Allows you to select joints"
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkObjects 
         Caption         =   "Select objects"
         Height          =   255
         Left            =   360
         TabIndex        =   50
         ToolTipText     =   "Allows you to select objects"
         Top             =   960
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDeselect 
         Caption         =   "Deselect"
         Height          =   375
         Left            =   360
         TabIndex        =   49
         ToolTipText     =   "Deselects the last object when a new object is clicked"
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Height          =   6615
      Index           =   1
      Left            =   7920
      TabIndex        =   8
      Top             =   -360
      Visible         =   0   'False
      Width           =   2900
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create Brush"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Click to create the brush"
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Cancel Brush"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         ToolTipText     =   "Click to cancel the selected options"
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.Slider ShpProp 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Sets the size of the top of the object"
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   20
         SelStart        =   20
         TickFrequency   =   2
         Value           =   20
      End
      Begin VB.OptionButton EditLine 
         Caption         =   "Move Profile"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "Moves the whole line that defines some objects"
         Top             =   4680
         Width           =   1695
      End
      Begin VB.OptionButton EditLine 
         Caption         =   "Extend line"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Alters the profile of the new object"
         Top             =   4440
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton EditLine 
         Caption         =   "Move Axis"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Sets the position of the axis"
         Top             =   4200
         Width           =   1695
      End
      Begin MSComctlLib.Slider ShpProp 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Sets the angle of the whole object"
         Top             =   2640
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Max             =   71
         TickFrequency   =   5
      End
      Begin MSComctlLib.Slider ShpProp 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Sets the number of edges"
         Top             =   2040
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   3
         Max             =   25
         SelStart        =   3
         Value           =   3
      End
      Begin VB.ListBox ShapeList 
         Height          =   1620
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2655
      End
      Begin MSComctlLib.Slider ShpProp 
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   25
         ToolTipText     =   "Sets the size of the bottom of the object"
         Top             =   3240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   20
         SelStart        =   20
         TickFrequency   =   2
         Value           =   20
      End
      Begin MSComctlLib.Slider ShpProp 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Sets the number of faces around the axis"
         Top             =   3840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   3
         Max             =   25
         SelStart        =   3
         Value           =   3
      End
      Begin VB.Label ShpName 
         Caption         =   "Horizontal faces"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   41
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label ShpName 
         Caption         =   "Bottom Face"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   27
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label ShpName 
         Caption         =   "Top face"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label ShpName 
         Caption         =   "Offset Angle"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label ShpName 
         Caption         =   "Edges"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame4"
      Height          =   5415
      Index           =   7
      Left            =   7320
      TabIndex        =   76
      Top             =   600
      Width           =   3015
      Begin VB.ListBox lstLights 
         Height          =   1230
         ItemData        =   "frmMain.frx":5425
         Left            =   360
         List            =   "frmMain.frx":542C
         TabIndex        =   108
         Tag             =   "0"
         Top             =   4080
         Width           =   2175
      End
      Begin MSComctlLib.Slider SLDQuality 
         Height          =   495
         Left            =   480
         TabIndex        =   103
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   4
         SelStart        =   4
         Value           =   4
      End
      Begin MSComctlLib.Slider sldLight 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   78
         ToolTipText     =   "The amount of light to pick out details"
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin MSComctlLib.Slider sldLight 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   77
         ToolTipText     =   "The amount of background light"
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.Label Label8 
         Caption         =   "Light Styles"
         Height          =   255
         Left            =   240
         TabIndex        =   109
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label lblRenderDis 
         Alignment       =   2  'Center
         Caption         =   "Wire frame - Outline only"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Render Quality"
         Height          =   255
         Left            =   480
         TabIndex        =   104
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Spot Light"
         Height          =   255
         Left            =   360
         TabIndex        =   80
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Ambiant Light"
         Height          =   255
         Left            =   360
         TabIndex        =   79
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame4"
      Height          =   4215
      Index           =   8
      Left            =   7920
      TabIndex        =   96
      Top             =   1560
      Width           =   2775
      Begin VB.CheckBox chk3DOp 
         Caption         =   "Label Joints"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   107
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chk3DOp 
         Caption         =   "Remove Hidden Faces"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   106
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chk3DOp 
         Caption         =   "Show Origin"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   100
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CheckBox chk3DOp 
         Caption         =   "Apply Perspective"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   99
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chk3DOp 
         Caption         =   "Draw Skeliton"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   98
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chk3DOp 
         Caption         =   "Draw Polygons"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   97
         Top             =   960
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Draw options"
         Height          =   255
         Left            =   360
         TabIndex        =   101
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Textures"
      Height          =   4695
      Index           =   0
      Left            =   6960
      TabIndex        =   93
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
      Begin VB.ListBox lstPaintTool 
         Height          =   780
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":5438
         Left            =   480
         List            =   "frmMain.frx":5442
         TabIndex        =   111
         Top             =   1800
         Width           =   1935
      End
      Begin VB.PictureBox picPallette 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2850
         Left            =   120
         Picture         =   "frmMain.frx":5451
         ScaleHeight     =   188
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   176
         TabIndex        =   102
         Top             =   2760
         Width           =   2670
      End
      Begin MSComctlLib.Slider LineSize 
         Height          =   255
         Left            =   360
         TabIndex        =   95
         ToolTipText     =   "Draw width"
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear texture"
         Height          =   495
         Left            =   240
         TabIndex        =   94
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Line width"
         Height          =   255
         Left            =   360
         TabIndex        =   110
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame1"
      Height          =   5655
      Index           =   4
      Left            =   3720
      TabIndex        =   44
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox RotateMove 
         Caption         =   "Check1"
         Height          =   255
         Left            =   360
         TabIndex        =   116
         Top             =   4560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optGetCenter 
         Caption         =   "Around vertex center"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   115
         Top             =   4200
         Width           =   2055
      End
      Begin VB.OptionButton optGetCenter 
         Caption         =   "Custom"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   85
         Top             =   3720
         Width           =   2295
      End
      Begin VB.OptionButton optGetCenter 
         Caption         =   "Around selection center"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   84
         Top             =   3240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optGetCenter 
         Caption         =   "Around world center"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   83
         Top             =   3480
         Width           =   1815
      End
      Begin VB.OptionButton optGetCenter 
         Caption         =   "Around joint"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   82
         Top             =   3960
         Width           =   1335
      End
      Begin VB.OptionButton optGetCenter 
         Caption         =   "Around object centers"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   81
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton QuickSpin 
         Height          =   495
         Index           =   3
         Left            =   1560
         Picture         =   "frmMain.frx":1D853
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Quick 90* rotate left"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton QuickSpin 
         Height          =   495
         Index           =   2
         Left            =   480
         Picture         =   "frmMain.frx":1DB5D
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Quick 90* rotate left"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton QuickSpin 
         Height          =   495
         Index           =   1
         Left            =   1560
         Picture         =   "frmMain.frx":1DE67
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Quick 5* rotate right"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton QuickSpin 
         Height          =   495
         Index           =   0
         Left            =   480
         Picture         =   "frmMain.frx":1E171
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Quick 5* rotate left"
         Top             =   1080
         Width           =   855
      End
      Begin MSComCtl2.UpDown udnChangeAngle 
         Height          =   375
         Left            =   2280
         TabIndex        =   47
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCustomAngle"
         BuddyDispid     =   196667
         OrigLeft        =   2280
         OrigTop         =   1920
         OrigRight       =   2520
         OrigBottom      =   2175
         Max             =   359
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCustomAngle 
         Height          =   375
         Left            =   1800
         TabIndex        =   46
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdRotate 
         Caption         =   "Rotate Brush"
         Height          =   375
         Left            =   360
         TabIndex        =   45
         ToolTipText     =   "Click to set the changes to your model"
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   360
         X2              =   2520
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   360
         X2              =   2520
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Rotate throught"
         Height          =   255
         Left            =   480
         TabIndex        =   72
         Top             =   520
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu OpFile 
         Caption         =   "&New"
         Index           =   1
      End
      Begin VB.Menu OpFile 
         Caption         =   "&Open..."
         Index           =   2
      End
      Begin VB.Menu OpFile 
         Caption         =   "&Save"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu OpFile 
         Caption         =   "S&ave as..."
         Index           =   4
      End
      Begin VB.Menu OpFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu OpFile 
         Caption         =   "&Import..."
         Index           =   6
      End
      Begin VB.Menu OpFile 
         Caption         =   "&Export..."
         Index           =   7
      End
      Begin VB.Menu OpFile 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu OpFile 
         Caption         =   "&Properties..."
         Index           =   9
      End
      Begin VB.Menu OpFile 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu OpFile 
         Caption         =   "E&xit"
         Index           =   11
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu OpEdit 
         Caption         =   "&Undo"
         Index           =   1
         Shortcut        =   ^Z
      End
      Begin VB.Menu OpEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu OpEdit 
         Caption         =   "&Copy"
         Index           =   3
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu OpEdit 
         Caption         =   "&Paste"
         Index           =   4
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu OpEdit 
         Caption         =   "Copy &to..."
         Index           =   5
      End
      Begin VB.Menu OpEdit 
         Caption         =   "Paste &from..."
         Index           =   6
      End
      Begin VB.Menu OpEdit 
         Caption         =   "&Duplicate"
         Index           =   7
         Shortcut        =   ^D
      End
      Begin VB.Menu OpEdit 
         Caption         =   "D&elete"
         Index           =   8
         Shortcut        =   {DEL}
      End
      Begin VB.Menu OpEdit 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu OpEdit 
         Caption         =   "Des&elect all"
         Index           =   10
      End
      Begin VB.Menu OpEdit 
         Caption         =   "&Select all"
         Index           =   11
      End
      Begin VB.Menu OpEdit 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu OpEdit 
         Caption         =   "&Group"
         Index           =   13
         Shortcut        =   ^G
      End
      Begin VB.Menu OpEdit 
         Caption         =   "U&ngroup"
         Index           =   14
      End
      Begin VB.Menu OpEdit 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu OpEdit 
         Caption         =   "Se&ttings..."
         Index           =   16
         Shortcut        =   ^{F1}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu SnapTo 
         Caption         =   "Snap to grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu Line 
         Caption         =   "-"
      End
      Begin VB.Menu opview 
         Caption         =   "&Zoom"
         Index           =   1
         Begin VB.Menu OpZoom 
            Caption         =   "&50%"
            Index           =   1
         End
         Begin VB.Menu OpZoom 
            Caption         =   "&100%"
            Checked         =   -1  'True
            Index           =   2
            Shortcut        =   {F9}
         End
         Begin VB.Menu OpZoom 
            Caption         =   "&200%"
            Index           =   3
         End
         Begin VB.Menu OpZoom 
            Caption         =   "&400%"
            Index           =   4
         End
         Begin VB.Menu OpZoom 
            Caption         =   "&800%"
            Index           =   5
         End
         Begin VB.Menu OpZoom 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu OpZoom 
            Caption         =   "&-25%"
            Index           =   7
            Shortcut        =   {F12}
         End
         Begin VB.Menu OpZoom 
            Caption         =   "&+25%"
            Index           =   8
            Shortcut        =   {F11}
         End
      End
      Begin VB.Menu opview 
         Caption         =   "Center View"
         Index           =   2
      End
      Begin VB.Menu opview 
         Caption         =   "Grayed out"
         Index           =   3
      End
      Begin VB.Menu opview 
         Caption         =   "&Rulers"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu opview 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu opview 
         Caption         =   "&Print Screen"
         Index           =   6
         Begin VB.Menu OpPrint 
            Caption         =   "&Print Screen"
            Index           =   1
            Shortcut        =   {F8}
         End
         Begin VB.Menu OpPrint 
            Caption         =   "&Record Scene..."
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuAnimation 
      Caption         =   "&Animation"
      Begin VB.Menu OpAnimation 
         Caption         =   "&Edit Scenes..."
         Index           =   1
      End
      Begin VB.Menu OpAnimation 
         Caption         =   "Animation Viewer..."
         Index           =   2
      End
      Begin VB.Menu OpAnimation 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu OpAnimation 
         Caption         =   "Scenes"
         Index           =   4
         Begin VB.Menu MnuScene 
            Caption         =   "No scene"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu OpHelp 
         Caption         =   "&About..."
         Index           =   1
      End
      Begin VB.Menu OpHelp 
         Caption         =   "&Help..."
         Index           =   2
      End
      Begin VB.Menu OpHelp 
         Caption         =   "&Customize toolbar"
         Index           =   3
         Begin VB.Menu OpCustomize 
            Caption         =   "&Main"
            Index           =   1
         End
         Begin VB.Menu OpCustomize 
            Caption         =   "&Edit"
            Index           =   2
         End
         Begin VB.Menu OpCustomize 
            Caption         =   "&Misq"
            Index           =   3
         End
         Begin VB.Menu OpCustomize 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu OpCustomize 
            Caption         =   "&Standard"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu OpCustomize 
            Caption         =   "&Flat"
            Index           =   6
         End
      End
      Begin VB.Menu OpHelp 
         Caption         =   "&Show notes"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   {F1}
      End
      Begin VB.Menu OpHelp 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu OpHelp 
         Caption         =   "&Edit Help..."
         Index           =   6
      End
   End
   Begin VB.Menu mnuJointPopup 
      Caption         =   "JointPopup"
      Visible         =   0   'False
      Begin VB.Menu JPopup 
         Caption         =   "Rename"
         Index           =   1
      End
      Begin VB.Menu JPopup 
         Caption         =   "Set as defalut"
         Index           =   2
      End
   End
   Begin VB.Menu mnuEditPopUp 
      Caption         =   "EditPopUp"
      Visible         =   0   'False
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Copy"
         Index           =   3
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Paste"
         Index           =   4
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Duplicate"
         Index           =   5
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Delete"
         Index           =   6
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Group"
         Index           =   8
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Ungroup"
         Index           =   9
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Object properties..."
         Index           =   11
      End
      Begin VB.Menu OpEditPopUp 
         Caption         =   "Link selected to..."
         Index           =   12
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MD3D As Byte         'holder for mouse down in 3d window
Dim Xstart As Integer       'position user first clicked on
Dim Ystart As Integer
Dim Angle1 As Single, Angle2 As Single, Angle3 As Single
Dim LightStyle(10) As String
Dim RecordScene As Boolean
Dim MoveLine As Byte, Scaleline As Byte
Dim DontDeselect As Boolean
Dim Rotation As Integer

Public Sub CenterView()
    If ViewMode < 4 Then
        Dim NowOn As Integer
        If CountSelectedObject = 0 Then Exit Sub
        For NowOn = 1 To cstTotalObjects
            If Object(NowOn).Selected = True Then
                X = X + Functions.FindCenter("X", NowOn)
                Y = Y + Functions.FindCenter("Y", NowOn)
                Z = Z + Functions.FindCenter("Z", NowOn)
            End If
        Next NowOn
        X = X / CountSelectedObject
        Y = Y / CountSelectedObject
        Z = Z / CountSelectedObject
        If ViewMode = 1 Then UDBar = Z: LRBar = X
        If ViewMode = 2 Then UDBar = Y: LRBar = X
        If ViewMode = 3 Then UDBar = Y: LRBar = Z
    Else
        Angle1 = 0
        Angle2 = 0
        Angle3 = 0
    End If
    frmMain.DrawModel
End Sub

Private Sub Block_Click()
    MsgBox "Magic Mushrooms rule the world!!!"
End Sub

Private Sub CBar_HeightChanged(ByVal NewHeight As Single)
    SortOutScreen
End Sub

Private Sub SetSideBar(SideBar)
    For N = 0 To Frame.Count - 1: Frame(N).Visible = False: Next N
    Frame(SideBar).Visible = True
End Sub

Private Sub chk3DOp_Click(Index As Integer)
    DrawModel
End Sub


Private Sub chkVert_Click()
    DrawModel
    
    If chkVert = 1 Then X = 0
    If chkVert = 0 Then X = 1
    
    optGetCenter(5).Enabled = chkVert
    optSeperate.Enabled = chkVert
    optDelVertecies.Enabled = chkVert
    
    edMode(6) = True
    
    chkSelDesel.Enabled = X
    chkDeselect.Enabled = X
    chkJoints.Enabled = X
    chkObjects.Enabled = X
    chkAttchOb.Enabled = X
    chkTotalSelect.Enabled = X
End Sub

Private Sub Command1_Click()
    Texture.Cls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    X = Chr$(KeyAscii)
    If DirectX.Visible = True Then
        Select Case X
            Case "8"
                MovePlayerForward (1)
            Case "5"
                MovePlayerForward (-1)
            Case "4"
                RotatePlayer (-4 / Pye)
            Case "6"
                RotatePlayer (4 / Pye)
            Case "1"
                MovePlayerUp (1)
            Case "0"
                MovePlayerUp (-1)
        End Select
        RenderScene
    End If
End Sub


Private Function LoadSettings() As Boolean
    On Error GoTo LoadFailed
    LoadSettings = False
    Open App.Path & "\data\Config.dat" For Input As #1
        Input #1, c: frmSettings.Clours(0).BackColor = c: Colours(3) = c
        Input #1, c: frmSettings.Clours(1).BackColor = c: Colours(4) = c
        Input #1, c: frmSettings.Clours(2).BackColor = c: Colours(5) = c
        Input #1, c: frmSettings.Clours(3).BackColor = c: Colours(6) = c
        Input #1, c: frmSettings.Clours(4).BackColor = c: Colours(7) = c
        
        Input #1, c: frmSettings.chkHilighBox = c
        Input #1, c: frmSettings.chkNew = c
        Input #1, c: frmSettings.DataSpin = c
        Input #1, c: frmSettings.GridSize = c
        Input #1, c: frmSettings.chkCenter = c
        Input #1, Index
        Input #1, c: frmSettings.sldGrey = c
        frmSettings.lblgx.BackColor = RGB(c * 10, c * 10, c * 10)
        If Index = 5 Or Index = 6 Then
            If Index = 5 Then TBar1.Style = 0: TBar2.Style = 0: TBar3.Style = 0
            If Index = 6 Then TBar1.Style = 1: TBar2.Style = 1: TBar3.Style = 1
            OpCustomize(5).Checked = False
            OpCustomize(6).Checked = False
            OpCustomize(Index).Checked = True
        End If
        
        frmMain.Axis.BorderColor = Colours(Index + 3)
    LoadSettings = True
LoadFailed:
Close
End Function

Private Sub Form_Load()
    Load frmSplash
    frmSplash.Visible = True
    Me.Visible = False
    picPallette.Tag = vbBlack
    frmSplash.Refresh
    InitRM DirectX
    CheckForFrames
    InitScene
    Make_LookUp
    DoEvents
    SideBar.ZOrder 1
    For N = 0 To Frame.Count - 1: Frame(N).Visible = False: Next N: Frame(5).Visible = True
    CurrentSideBar = 5
    ViewMode = 1
    Zoom = 1
    If LoadSettings = False Then MsgBox "Couldn't load program settings", vbInformation, "File not found"
    Model.Saved = True
    lstPaintTool.Selected(0) = True
    lstLights.AddItem "Glow": LightStyle(1) = "LLLLMMMMNNNOOOPPPQQRRSTUVWWXXYYYZZZZZZZZZYYYXXWWVUTSRRQQPPPOOONNNNMMMM"
    lstLights.AddItem "Bright Flicker": LightStyle(2) = "QQQQQQRSRQQQQQQRTSTRQQQQQSRTQQQQQQQQSRTZZZZZZEQQQQ"
    lstLights.AddItem "Lightning": LightStyle(3) = "AAAAAAAAZZAAAAAAAAAAAAAAAAZAAAAAAAAAAAAAZAZAZAAAAAAAAAAAAAAGHJBZZZZHJZZHAAAAAAAAAAAAAAAAAAZZAGHGHGJZZZZZAAAAAAAAAAAAAA"
    lstLights.AddItem "Strobe": LightStyle(4) = "AAZZ"
    lstLights.AddItem "Flicker": LightStyle(5) = "BFJKLDEHFQJQWPBHCDQBWJKLYDJKLQDYKLPQWDYCKWLDHQHDJK"
    lstLights.Selected(0) = True
    NewShapeList
    SetNewShapeMenu
    NewModel 0
    Comand = Command
    'Comand = App.Path & "man2.am5"
    If Comand <> "" Then
        If LCase(Right(Comand, 3)) = "am5" Then
            If LCase(Left(Comand, 2)) = "/c" Then
                Comand = Mid(Comand, 4, 9999)
                frmMain.Tag = "Nooo"
                If LoadFile(Comand) = False Then
                    MsgBox "The file '" & Comand & "' could not be opened", vbCritical, "Error"
                Else
                    XX = Mid(Model.ProjectFileName, 1, Len(Model.ProjectFileName) - 4) & ".dat"
                    FrmExport.CompilerControler XX, 0
                End If
                Load FrmExport
                Unload frmSplash
                Destroy "shapes.txt"
                Unload Timei
                Unload Me
                End
            Else
                LoadFile Comand
            End If
        ElseIf LCase(Right(Comand, 3)) = "abf" Then
            ImportStuff.LoadBatchFile Command
        Else
            MsgBox "Unrecognised file format", vbCritical, "Error"
        End If
    End If
    DrawModel
    Me.Visible = True
    PresetScale = Scales.Text
    Unload frmSplash
    SBar.Panels(3) = "Welcome to " & App.Title & " " & App.Major & ".   Look in the help menu if you are stuck!"
End Sub

Private Sub Form_Resize()
    SortOutScreen
    DrawModel
End Sub

Public Sub SortOutScreen()
    On Error Resume Next
    SideBar.Width = 3000
    If Me.ScaleWidth < 4000 Then SideBar.Width = Me.ScaleWidth - 1000
    CBar.Width = Me.ScaleWidth:    MainTab.Width = Me.ScaleWidth - SideBar.Width - 50
    MainTab.Height = Me.ScaleHeight - CBar.Height - SBar.Height - 50
    MainTab.Top = CBar.Height
    If TBar3.buttons(7).Value = tbrPressed Then MainTab.Width = Me.ScaleWidth
    SideBar.Left = Me.ScaleWidth - SideBar.Width
    SideBar.Top = MainTab.ClientTop - 30
    SideBar.Height = MainTab.ClientHeight + 80
    view.Top = MainTab.ClientTop:    view.Left = MainTab.ClientLeft
    view.Width = MainTab.ClientWidth:    view.Height = MainTab.ClientHeight
    DirectX.Top = view.Top:    DirectX.Left = view.Left
    DirectX.Width = view.Width:    DirectX.Height = view.Height
    TopRule.Left = Corner.Width:    TopRule.Width = view.ScaleWidth - Corner.Width
    SideRule.Top = Corner.Height:    SideRule.Height = view.Height - Corner.Height
    For N = 0 To Frame.Count
        With Frame(N)
            .Top = SideBar.Top + 50
            .Left = SideBar.Left + 50
            .Width = SideBar.Width - 150
            .Height = SideBar.Height - 150
            .BorderStyle = 0
        End With
    Next N
    UDBar.Left = view.ScaleWidth - UDBar.Width:    LRBar.Top = view.ScaleHeight - LRBar.Height
    UDBar.Height = view.ScaleHeight - LRBar.Height:    LRBar.Width = view.ScaleWidth - UDBar.Width
    Block.Left = UDBar.Left:    Block.Top = LRBar.Top:    ShowTime.Width = SBar.Panels(2).Width - 60
    ShowTime.Left = SBar.Panels(2).Left + 30:    ShowTime.Height = SBar.Height - 80
    ShowTime.Top = Me.ScaleHeight - ShowTime.Height - 20
    cmdCreate(0).Top = Frame(1).Height - cmdCreate(0).Height - 200
    cmdCreate(1).Top = Frame(1).Height - cmdCreate(1).Height - 200
    cmdScale.Top = Frame(2).Height - cmdScale.Height - 200
    cmdRotate.Top = Frame(2).Height - cmdRotate.Height - 200
    cmdEdit.Top = Frame(2).Height - cmdRotate.Height - 200
    Joints.Height = Frame(3).Height - 2000:    SkeFrame.Top = Joints.Height + 350
    opSelectJ.Top = Joints.Height + 400
    opAddJ.Top = Joints.Height + 800
    opChangeJ.Top = Joints.Height + 1200
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If frmMain.Tag <> "Nooo" Then
        Cancel = 1
        If YouWantToQuit = True Then End
    End If
End Sub

Private Sub MainTab_Click()
    ViewMode = MainTab.SelectedItem.Index
    DirectX.Visible = False
    Texture.Visible = False
    TBar3.buttons(4).Enabled = False
    TBar3.buttons(5).Enabled = False
    TopRule.Visible = True
    SideRule.Visible = True
    Corner.Visible = True
    Select Case MainTab.SelectedItem.Index
        Case 1 To 3
            UDBar.Visible = True
            LRBar.Visible = True
            Block.Visible = True
            CBar.Bands(3).Visible = True
            SetSideBar CurrentSideBar
            TBar3.buttons(4).Value = tbrPressed
            StopAnimation
            If frmSettings.chkCenter = 1 Then CenterView
        Case 4
            CBar.Bands(3).Visible = False
            Texture.Visible = False
            SetSideBar 8
            TBar3.buttons(4).Enabled = True
            TBar3.buttons(5).Enabled = True
            TopRule.Visible = False
            SideRule.Visible = False
            Corner.Visible = False
            UDBar.Visible = False
            LRBar.Visible = False
            Block.Visible = False
        Case 5
            CBar.Bands(3).Visible = False
            DirectX.Visible = True
            SetSideBar 7
            DirectXh.PlaceModelinWindow
            RenderScene
            StopAnimation
        Case 6
            CBar.Bands(3).Visible = False
            SetSideBar 0
            Texture.Visible = True
            StopAnimation
    End Select
    DrawModel
End Sub

Private Sub mnuHelp_Click()
    If Model.ModelNotes = "" Then OpHelp(4).Enabled = False Else OpHelp(4).Enabled = True
End Sub

Private Sub OpAnimation_Click(Index As Integer)
    Select Case Index
        Case 1
            If CountJoints = 0 Then
                MsgBox "This animation method relies on joints and skelitons" & vbNewLine & "To create a skeliton, use the last tool on the toolbar, or have a look in help", vbInformation
                Exit Sub
            End If
            sceneEdit.Start
            sceneEdit.Visible = True
            frmMain.Visible = False
        Case 2
            If CountObjects = 0 Or CountJoints = 0 Then
                frmView.Visible = True
                frmView.Start ""
                frmMain.Visible = False
                Exit Sub
            End If
            X = "For speed, this feature only works on compiled models, which means" & vbNewLine
            X = X & "you must compile this model before you view it." & vbNewLine & vbNewLine
            X = X & "Compiling usually takes well under a minute, depending on how" & vbNewLine
            X = X & "large it is. Do you want to compile it now?" & vbNewLine
            Reponce = MsgBox(X, vbYesNo + vbQuestion, "Do you want to...")
            If Reponce = 6 Then
                XX = App.Path & "\data\t.dat"
                FrmExport.CompilerControler XX, 0
                frmView.Visible = True
                frmView.Start App.Path & "\data\t.dat"
            Else
                frmView.Visible = True
                frmView.Start ""
            End If
            frmMain.Enabled = False
    End Select
End Sub

Private Sub OpEdit_Click(Index As Integer)
    Select Case Index
        Case 1: UndoLastMove: DrawModel
        Case 2:
        Case 3:
            If Functions.CountSelectedObject = 0 Then frmMain.SBar.Panels(3) = "###### Error - Nothing to copy ########": Exit Sub
            CopyTo App.Path & "\copyfile": DrawModel
        Case 4: DeselectAll: PasteFrom App.Path & "\copyfile", 0: DrawModel
        Case 5
            FileName = SelectFileName("Copy", "Enter filename to copy to...")
            If FileName <> "" Then CopyTo FileName
        Case 6:
            FileName = SelectFileName("Copy", "Select filename to paste from...")
            If FileName <> "" Then PasteFrom FileName, 1
        Case 7: CopyTo App.Path & "\Duplicatefile": DeselectAll: PasteFrom App.Path & "\Duplicatefile", 0: On Error Resume Next: Kill App.Path & "\Duplicatefile": DrawModel
        Case 8:
            frmMain.OpEdit(1).Caption = "Undo delete"
            StoreCurentPosition
            frmMain.OpEdit(1).Enabled = True
            DeleteSelected
            DrawModel
        Case 10:
            DeselectAll
            DrawModel
        Case 11:
            SelectAll
            DrawModel
        Case 13: GroupSelected: DrawModel
        Case 14: UnGroupSelected: DrawModel
        Case 16: frmSettings.Visible = True: Me.Enabled = False
    End Select
End Sub

Private Sub OpFile_Click(Index As Integer)
    Select Case Index
        Case 1:
            NewModel 1
        Case 2:
            FileName = SelectFileName("Am5", "Open file...")
            If FileName <> "" Then
                If LCase(Right(FileName, 3)) = "am5" Then
                    LoadFile FileName
                Else
                    LoadBatchFile FileName
                End If
            End If
        Case 3:
            If Model.ProjectFileName <> "*.txt" Then
                SaveFile Model.ProjectFileName
            End If
        Case 4:
            FileName = SetFileName("Am5", "Save file...")
            If FileName <> "" Then
                If CheckOverwrite(FileName) = False Then Exit Sub
                SaveFile FileName
            End If
        Case 5:
        Case 6:
            FileName = SelectFileName("Import", "Import file...")
            If FileName <> "" Then ImportModel FileName
        Case 7: FrmExport.Visible = True: Me.Enabled = False
        Case 8:
        Case 9: frmProperties.Visible = True: Me.Enabled = False
        Case 11: If YouWantToQuit = True Then End
    End Select
End Sub

Private Sub OpHelp_Click(Index As Integer)
    Select Case Index
        Case 1: frmAbout.Visible = True: Me.Enabled = False
        Case 2
            Select Case CurrentSideBar
                Case 5: ShowHelp "Main"
                Case 1: ShowHelp "Creating Objects"
                Case 6: ShowHelp "Deforming objects"
                Case 4: ShowHelp "Rotating Objects"
                Case 2: ShowHelp "Scaling Objects"
                Case 3: ShowHelp "Creating a Skeliton"
            
                
            End Select
        Case 4
            AlwaysOnTop frmShowMessage, 1
            frmShowMessage.txtNotes = Model.ModelNotes
            frmShowMessage.Caption = Model.ProjectName
            frmShowMessage.Visible = True
            frmShowMessage.SetFocus
        Case 6
            X = InputBox("Enter password, please...")
            If X = "QWERT" Then
                If ShowHelp("EditMeNow!!!") = False Then MsgBox "There was an error staring the help program", , "Error"
            End If
    End Select
End Sub

Private Sub optDelVertecies_Click()
    cmdEdit.Enabled = True
    cmdEdit.Caption = "Delete Vertecies"
End Sub

Private Sub optJoint_Click()
    cmdEdit.Enabled = True
    cmdEdit.Caption = "Joint objects"
End Sub

Private Sub optSeperate_Click()
    cmdEdit.Enabled = True
    cmdEdit.Caption = "Seperate object"
End Sub

Private Sub OpZoom_Click(Index As Integer)
    For N = 1 To 8: OpZoom(N).Checked = False: Next
    If Index < 7 Then OpZoom(Index).Checked = True
    Select Case Index
        Case 1: Zoom = 0.5
        Case 2: Zoom = 1
        Case 3: Zoom = 2
        Case 4: Zoom = 4
        Case 5: Zoom = 8
        Case 7: Zoom = Zoom - 0.25
        Case 8: Zoom = Zoom + 0.25
    End Select
    If Zoom > 8 Then Zoom = 8
    If Zoom < 0.25 Then Zoom = 0.25
    SBar.Panels(3) = "Zoom = " & Zoom * 100 & "%"
    DrawModel
End Sub

Private Sub OpEditPopUp_Click(Index As Integer)
    Select Case Index
        Case 1: UndoLastMove: DrawModel
        Case 3:
            If Functions.CountSelectedObject = 0 Then frmMain.SBar.Panels(3) = "###### Error - Nothing to copy ########": Exit Sub
            CopyTo App.Path & "\copyfile": DrawModel
        Case 4: DeselectAll: PasteFrom App.Path & "\copyfile", 0: DrawModel
        Case 5: CopyTo App.Path & "\Duplicatefile": DeselectAll: PasteFrom App.Path & "\Duplicatefile", 0: On Error Resume Next: Kill App.Path & "\Duplicatefile": DrawModel
        Case 6
            frmMain.OpEdit(1).Caption = "Undo delete"
            StoreCurentPosition
            frmMain.OpEdit(1).Enabled = True
            DeleteSelected
            DrawModel
        
        Case 8: GroupSelected: DrawModel
        Case 9: UnGroupSelected: DrawModel
        Case 11: EditObjectProperties
        Case 12
            JointOn = JointOver(Xstart, Ystart)
            If JointOn = 0 Then Beep: SBar.Panels(3) = "####### Error - Couldn't find joint to link to! #######": Exit Sub
            For N = 1 To cstTotalObjects
                If Object(N).Selected = True Then
                    For M = 1 To Object(N).VertexCount
                        If CurrentSideBar = 5 And chkVert = 1 Then
                            If Object(N).Vertex(M).Selected = True Then
                                Object(N).Vertex(M).TargetName = BaseFrame(JointOn).Key
                            End If
                        Else
                            Object(N).Vertex(M).TargetName = BaseFrame(JointOn).Key
                        End If
                    Next M
                End If
            Next N
            SBar.Panels(3) = "Objects linked to " & BaseFrame(JointOn).Name
    End Select
End Sub



Private Sub picPallette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picPallette.Cls
        picPallette.Circle (X, Y), 4
        picPallette.Tag = picPallette.Point(X, Y)
    End If
End Sub

Private Sub picPallette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picPallette.Cls
        picPallette.Circle (X, Y), 4
        picPallette.Tag = picPallette.Point(X, Y)
    End If
End Sub

Private Sub SLDQuality_Click()
    Select Case SLDQuality
        Case 1: lblRenderDis = "Points - Show vertecies only"
        Case 2: lblRenderDis = "Wire frame - Outline only"
        Case 3: lblRenderDis = "Flat Shade - One colour per face"
        Case 4: lblRenderDis = "Groud - Smooth blended faces"
    End Select
    SetMode SLDQuality: RenderScene
End Sub

Private Sub SLDQuality_Scroll()
    Select Case SLDQuality
        Case 1: lblRenderDis = "Points - Show vertecies only"
        Case 2: lblRenderDis = "Wire frame - Outline only"
        Case 3: lblRenderDis = "Flat Shade - One colour per face"
        Case 4: lblRenderDis = "Groud - Smooth blended faces"
    End Select
    SetMode SLDQuality: RenderScene
End Sub



Private Sub SnapTo_Click()
    If SnapTo.Checked = True Then
        SnapTo.Checked = False
    Else
        SnapTo.Checked = True
    End If
End Sub

Private Sub TBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            NewModel 1
        Case 2
            FileName = SelectFileName("Am5", "Open file...")
            If FileName <> "" Then
                If LCase(Right(FileName, 3)) = "am5" Then
                    LoadFile FileName
                Else
                    LoadBatchFile FileName
                End If
            End If
        Case 3:
            If Model.ProjectFileName <> "*.txt" Then
                SaveFile Model.ProjectFileName
            End If
        Case 5:
            If Functions.CountSelectedObject = 0 Then frmMain.SBar.Panels(3) = "###### Error - Nothing to copy ########": Exit Sub
            CopyTo App.Path & "\copyfile": DrawModel
        Case 6:
            PasteFrom App.Path & "\copyfile", 1: DrawModel
        Case 7:
            If Functions.CountSelectedObject = 0 Then frmMain.SBar.Panels(3) = "###### Error - Nothing to copy ########": Exit Sub
            CopyTo App.Path & "\copyfile": DeleteSelected: DrawModel
    End Select
End Sub

Private Sub TBar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: CurrentSideBar = 5: DrawModel
        Case 2: CurrentSideBar = 1
        Case 3: CurrentSideBar = 6: OpEdit(1).Caption = "Undo edit": StoreCurentPosition: frmMain.OpEdit(1).Enabled = True: DrawModel
        Case 4: CurrentSideBar = 2: OpEdit(1).Caption = "Undo scale": StoreCurentPosition: frmMain.OpEdit(1).Enabled = True
        Case 5: CurrentSideBar = 4: OpEdit(1).Caption = "Undo rotate": StoreCurentPosition: frmMain.OpEdit(1).Enabled = True
        Case 6: CurrentSideBar = 3
    End Select
End Sub

Private Sub TBar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Zoom = Zoom - 0.25: If Zoom < 0.25 Then Zoom = 0.25
            For N = 1 To 8: OpZoom(N).Checked = False: Next
            SBar.Panels(3) = "Zoom = " & Zoom * 100 & "%"
            DrawModel
        Case 2
            Zoom = Zoom + 0.25: If Zoom > 8 Then Zoom = 8
            For N = 1 To 8: OpZoom(N).Checked = False: Next
            SBar.Panels(3) = "Zoom = " & Zoom * 100 & "%"
            DrawModel
        Case 3:
        Case 4:
            StopAnimation
        Case 5:
             StartAnimation
        Case 7
            SortOutScreen
            DrawModel
    End Select
End Sub

Public Sub StopAnimation()
    Animate = False
    frmMain.Label3 = ""
    DrawModel
    TBar3.buttons(4).Value = tbrPressed
End Sub

Public Sub DrawModel()
    DirectX.Visible = False
    view.AutoRedraw = True
    frmMain.view.Cls
    view.DrawStyle = 0
    If ViewMode < 4 Then DrawRulers
    If ViewMode = 1 Then DrawFromTop
    If ViewMode = 2 Then DrawFromSide
    If ViewMode = 3 Then DrawFromFront
    If ViewMode = 4 Then
        FindNewSkeliton
        If chk3DOp(3) = 1 Then DrawGuides view, Angle1, Angle2, Angle3, view.ScaleWidth / 2, view.ScaleHeight / 2
        If chk3DOp(1) = 1 Then Draw3DSkeliton view, Angle1, Angle2, Angle3, view.ScaleWidth / 2, view.ScaleHeight / 2
        If chk3DOp(0) = 1 Then
            For N = 1 To cstTotalObjects
                If Object(N).Used = True Then
                    ThreeDEngine.Draw3DBrush frmMain.view, N, Angle1, Angle2, Angle3, 0, 0, 0, 0, 0, 0
                End If
            Next N
        End If
    End If
    If ViewMode = 5 Then
    End If
    If ViewMode = 5 Then
        DirectX.Visible = True
        'AnimateDirectX
    '    PlaceModelinWindow
      '  SetLights sldLight(0), sldLight(1)
    End If
    If ViewMode = 6 Then
        Texture.Top = -UDBar
        Texture.Left = -LRBar
        view.Circle (Texture.Left + Texture.Width, Texture.Top + Texture.Height), 5
    End If
End Sub

Private Sub UDBar_Change()
    DrawModel
End Sub

Private Sub UDBar_Scroll()
    DrawModel
    view.Refresh
End Sub

Private Sub LRBar_Change()
    DrawRulers
    DrawModel
End Sub

Private Sub LRBar_Scroll()
    DrawRulers
    DrawModel
    view.Refresh
End Sub

Private Sub cmdCreate_Click(Index As Integer)
    If Index = 0 Then
        createobject
    Else
        Guide.Tag = False
        LineLength = 0
        Guide.Visible = False
        DrawModel
    End If
End Sub

Private Sub cmdEdit_Click()
    If optJoint = True Then
        OpEdit(1).Caption = "Undo join objects"
        OpEdit(1).Enabled = True
        OpEditPopUp(1).Enabled = True
        First = FirstObject

        For N = First + 1 To cstTotalJoints
            If Object(N).Selected = True Then
                ReDim Preserve Object(First).Vertex(Object(First).VertexCount + Object(N).VertexCount) As VertexDis
                For M = 1 To Object(N).VertexCount
                    Object(First).Vertex(Object(First).VertexCount + M) = Object(N).Vertex(M)
                Next M
                ReDim Preserve Object(First).Face(Object(First).EdgeCount + Object(N).EdgeCount) As Integer
                EdgeOn = 0
                For M = 1 To Object(N).FaceCount
                    EdgeOn = EdgeOn + 1
                    Faces = Object(N).Face(EdgeOn)
                    Object(First).Face((Object(First).EdgeCount + EdgeOn)) = Object(N).Face(EdgeOn)
                    For o = 1 To Faces
                        EdgeOn = EdgeOn + 1
                        Object(First).Face((Object(First).EdgeCount + EdgeOn)) = Object(N).Face(EdgeOn) + Object(First).VertexCount
                    Next o
                Next M
                Object(First).VertexCount = Object(First).VertexCount + Object(N).VertexCount
                Object(First).EdgeCount = Object(First).EdgeCount + Object(N).EdgeCount
                Object(First).FaceCount = Object(First).FaceCount + Object(N).FaceCount
            End If
        Next N
        Object(First).Selected = False: DeleteSelected
        Object(First).Selected = True: FindOutLine First
        FindObjectOutline
        DrawModel


        Exit Sub
    End If

    If optSeperate = True Or optDelVertecies = True Then
        If optSeperate = True And CountSelectedObject <> 1 Then Exit Sub
        For N = 1 To cstTotalObjects
            If Object(N).Used = True And Object(N).Selected = True Then
                If optSeperate = True Then
                    Dim TempFace() As Integer, AddThis As Integer
                    Add = AddObject
                    Object(Add).Used = True
                    Object(Add).VertexCount = 0
                    Object(Add).FaceCount = 0
                    Object(Add).EdgeCount = 0
                    ReDim Object(Add).Face(0) As Integer
                    Object(Add).Colour = Object(N).Colour
                    For M = 1 To Object(N).VertexCount
                        If Object(N).Vertex(M).Selected = True Then
                            Object(Add).VertexCount = Object(Add).VertexCount + 1
                            ReDim Preserve Object(Add).Vertex(Object(Add).VertexCount) As VertexDis
                            Object(Add).Vertex(Object(Add).VertexCount) = Object(N).Vertex(M)
                            Object(Add).Vertex(Object(Add).VertexCount).Selected = False
                        End If
                    Next M
                    EdgeOn = 1
                    For M = 1 To Object(N).FaceCount
                       FaceCount = Object(N).Face(EdgeOn): EdgeOn = EdgeOn + 1
                       ReDim TempFace(FaceCount + 1) As Integer
                       TempFace(1) = FaceCount
                       AddThis = 0
                       For L = 1 To FaceCount
                           Vert = Object(N).Face(EdgeOn)
                           TempFace(L + 1) = Vert
                           EdgeOn = EdgeOn + 1
                           If Object(N).Vertex(Vert).Selected = True Then
                               AddThis = AddThis + 1
                           End If
                       Next L
                       
                       
                       If AddThis = FaceCount Then
                           
                           
                           Object(Add).EdgeCount = Object(Add).EdgeCount + FaceCount + 1
                           Object(Add).FaceCount = Object(Add).FaceCount + 1
                           ReDim Preserve Object(Add).Face(Object(Add).EdgeCount) As Integer
                           Object(Add).Face(Object(Add).EdgeCount - FaceCount) = TempFace(1)
                           For L = 2 To FaceCount + 1
                               For g = 1 To Object(Add).VertexCount
                                    If Object(Add).Vertex(g).X = Object(N).Vertex(TempFace(L)).X Then
                                        If Object(Add).Vertex(g).Y = Object(N).Vertex(TempFace(L)).Y Then
                                            If Object(Add).Vertex(g).Z = Object(N).Vertex(TempFace(L)).Z Then
                                                Object(Add).Face(Object(Add).EdgeCount - FaceCount + L - 1) = g
                                            End If
                                        End If
                                    End If
                               Next g
                           Next L
                       End If
                       
                   Next M
                   FindOutLine Add
                   FindObjectOutline
                   Me.chkVert.Value = 0
                   Object(Add).Selected = True
        
                End If



RestartDel:
                EdgeOn = 1
                For M = 1 To Object(N).FaceCount
                    FaceCount = Object(N).Face(EdgeOn): EdgeOn = EdgeOn + 1
                    For L = 1 To FaceCount
                        Vert = Object(N).Face(EdgeOn)
                        EdgeOn = EdgeOn + 1
                        If Object(N).Vertex(Vert).Selected = True Then
                            DeleteFace N, M
                            GoTo RestartDel
                        End If
                    Next L
                Next M
            End If
RestartVertDel:
            For M = 1 To Object(N).VertexCount
                If Object(N).Vertex(M).Selected = True Then
                    DeleteVertex N, M
                    GoTo RestartVertDel
                End If
            Next M
            FindOutLine (N)
        Next N
        FindObjectOutline
        DrawModel
        
        
        
        
        Exit Sub
    End If
    
    
    
    If optConvert = True Then
        OpEdit(1).Caption = "Undo convert faces"
        OpEdit(1).Enabled = True
        OpEditPopUp(1).Enabled = True
        
        For N = 1 To cstTotalObjects
            If Object(N).Selected = True Then
                ConvertThisToTriangles (N)
            End If
        Next N
        DrawModel
        Exit Sub
    End If
    
    Counta = 0
    If optInvert = True Then
        For N = 1 To cstTotalObjects
            If Object(N).Selected = True Then
                FlipFaces N
                Counta = Counta + 1
            End If
        Next N
        SBar.Panels(3) = Counta & " objects inverted"
        Exit Sub
    End If
    If optFlip(0) = True Then flipon = 0
    If optFlip(1) = True Then flipon = 1
    Dim NowOn As Integer
    Cx = FindCenter("X", 0)
    Cy = FindCenter("y", 0)
    Cz = FindCenter("z", 0)
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Selected = True Then
            Counta = Counta + 1
            FlipFaces NowOn
            For N = 1 To Object(NowOn).VertexCount
                Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X - Cx
                Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y - Cy
                Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z - Cz
            Next N
            For N = 1 To Object(NowOn).VertexCount
                If ViewMode = 1 Then
                    If flipon = 1 Then Object(NowOn).Vertex(N).Z = -Object(NowOn).Vertex(N).Z
                    If flipon = 0 Then Object(NowOn).Vertex(N).X = -Object(NowOn).Vertex(N).X
                End If
                If ViewMode = 2 Then
                    If flipon = 1 Then Object(NowOn).Vertex(N).Y = -Object(NowOn).Vertex(N).Y
                    If flipon = 0 Then Object(NowOn).Vertex(N).X = -Object(NowOn).Vertex(N).X
                End If
                If ViewMode = 3 Then
                    If flipon = 1 Then Object(NowOn).Vertex(N).Y = -Object(NowOn).Vertex(N).Y
                    If flipon = 0 Then Object(NowOn).Vertex(N).Z = -Object(NowOn).Vertex(N).Z
                End If
            Next N
            For N = 1 To Object(NowOn).VertexCount
                Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X + Cx
                Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y + Cy
                Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z + Cz
            Next N
        End If
        SelectDeselect.FindOutLine NowOn
    Next NowOn
    DrawModel
    SBar.Panels(3) = Counta & " objects fliped"
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If SnapTo.Checked = True Then
        If Index = 1 Then MoveSelected -frmSettings.GridSize, 0
        If Index = 2 Then MoveSelected 0, frmSettings.GridSize
        If Index = 3 Then MoveSelected frmSettings.GridSize, 0
        If Index = 0 Then MoveSelected 0, -frmSettings.GridSize
    Else
        If Index = 1 Then MoveSelected -1, 0
        If Index = 2 Then MoveSelected 0, 1
        If Index = 3 Then MoveSelected 1, 0
        If Index = 0 Then MoveSelected 0, -1
    End If
    DrawModel
End Sub

Private Sub cmdRotate_Click()
    If IsThisValid(txtCustomAngle, 1) = False Then
        MsgBox "Thats not a valid number (You can't have any letters)"
        Exit Sub
    End If
    SpinSelected txtCustomAngle
End Sub

Private Sub cmdScale_Click()
    If IsThisValid(PresetScale, 0) = False Then
        MsgBox "Thats not a valid number (You can't have '0' or any letters)"
        Exit Sub
    End If
    If SklMode(1) = True Then
        ScaleObject sclXDim(0) / 100, sclXDim(1) / 100, sclXDim(2) / 100, 0
    End If
    If SklMode(2) = True Then
        ScaleObject PresetScale, PresetScale, PresetScale, 0
    End If

    frmMain.DrawModel
End Sub

Private Sub CurrentSideBar_Change()
    For N = 1 To Frame.Count - 1
        Frame(N).Visible = False
    Next N
    If CurrentSideBar = 7 Then Frame(7).Visible = True
    If ViewMode > 3 Then Exit Sub
    If CurrentSideBar <> 0 Then Frame(CurrentSideBar).Visible = True
    SBar.Panels(3) = ""
    Select Case CurrentSideBar
        Case 1: SBar.Panels(3) = "Design and create new brushes"
        Case 2: SBar.Panels(3) = "Scale the selected objects"
        Case 3: SBar.Panels(3) = "View joints"
        Case 4: SBar.Panels(3) = "Move, rotate or flip the selected objects"
        Case 5: SBar.Panels(3) = "Select and move objects or joints"
        Case 6: SBar.Panels(3) = "Flip, reverse or change the shape of the selected objects"
    End Select
End Sub

Private Sub EditLine_Click(Index As Integer)
    If ShapeList.Text = "Torous" Then Exit Sub
    Guide.Visible = False
    If Index = 1 And LineLength <> 0 Then Guide.Visible = True: Guide.x1 = StoreLine(LineLength).X: Guide.y1 = StoreLine(LineLength).Y: Axis.BorderColor = 65536 - Colours(4)

End Sub

Private Sub EdMode_Click(Index As Integer)
    cmdEdit.Enabled = False
    If Index = 3 Then
        DeselectAllVertecies
    End If
    DrawModel
End Sub

Private Sub MnuScene_Click(Index As Integer)
    For N = 1 To MnuScene.Count - 1
        MnuScene(N).Checked = False
    Next N
    MnuScene(Index).Checked = True
    If Animate = True Then StartAnimation
End Sub

Private Sub optConvert_Click()
    cmdEdit.Enabled = True
    cmdEdit.Caption = "Convert faces"
End Sub

Private Sub optFlip_Click(Index As Integer)
    cmdEdit.Enabled = True
    cmdEdit.Caption = "Flip Brush"
End Sub

Private Sub optInvert_Click()
    cmdEdit.Enabled = True
    cmdEdit.Caption = "Invert Brush"
End Sub


Private Sub Joints_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuJointPopup
End Sub


Private Sub Joints_NodeClick(ByVal Node As MSComctlLib.Node)
    NodeOn = Val(Mid(Node.Key, 6, 4))
    For N = 1 To cstTotalJoints
     BaseFrame(N).Selected = False
    Next N
    BaseFrame(NodeOn).Selected = True
    DrawModel
End Sub

Private Sub JPopup_Click(Index As Integer)
    If Index = 1 Then
        Select Case Joints.SelectedItem.Tag
            Case 1
                X = InputBox("Enter a new name for your model", "Change name", Joints.SelectedItem.Text)
                If X <> "" Then
                    frmMain.Joints.Nodes("Model").Text = X
                    sceneEdit.Frame.Nodes("Model").Text = X
                    frmMain.Caption = App.Title & " [" & Model.ProjectName & "]"
                    Model.ProjectName = X
                    Block.Enabled = False
                    Model.Saved = False
                    If LCase(X) = "mushroom" Then Block.Enabled = True
                End If
            Case 2
                X = InputBox("Enter a new name for this joint", "Change name", Joints.SelectedItem.Text)
                If X <> "" Then
                    Joints.Nodes(Joints.SelectedItem.Key).Text = X
                    JointOn = Val(Mid(Joints.SelectedItem.Key, 6, 5))
                    BaseFrame(JointOn).Name = X
                    Block.Enabled = False
                    If LCase(X) = "mushroom" Then Block.Enabled = True
                    Model.Saved = False
                End If
        End Select
    End If
    If Index = 2 Then
        Select Case Joints.SelectedItem.Tag
            Case 2
                CurrentJoint = Joints.SelectedItem.Key
        End Select
    End If
End Sub



Private Sub OpCustomize_Click(Index As Integer)
    If Index = 1 Then TBar1.Customize
    If Index = 2 Then TBar2.Customize
    If Index = 3 Then TBar3.Customize
    If Index = 5 Or Index = 6 Then
        If Index = 5 Then TBar1.Style = 0: TBar2.Style = 0: TBar3.Style = 0
        If Index = 6 Then TBar1.Style = 1: TBar2.Style = 1: TBar3.Style = 1
        OpCustomize(5).Checked = False
        OpCustomize(6).Checked = False
        OpCustomize(Index).Checked = True
    End If
End Sub


Private Sub OpPrint_Click(Index As Integer)

If Index = 3 Then
    frmRecord.Visible = True
    Me.Enabled = False
    Do: DoEvents: Loop While frmRecord.Visible = True
    Me.Enabled = True
    Me.SetFocus
    If frmRecord.Tag = "Cancel" Then Exit Sub
    Rotation = frmRecord.Angle * frmRecord.Rotations
    MainTab.Enabled = False
    view.Enabled = False
    Exit Sub
End If

If Index = 2 And Animate = True Then
    MsgBox "To use this feature, Play mode must be switched off. See help for more details", , "Press the off switch!!"
    Exit Sub
End If

On Error GoTo CanceledPrintScreen
If Index = 1 Then GetFile.DialogTitle = "Select bitmap to save image to..."
If Index = 2 Then GetFile.DialogTitle = "Select bitmap to save scene to..."
frmMain.GetFile.Filter = "Bitmap image (*.bmp) |*.bmp"
GetFile.ShowSave
If Mid(GetFile.FileName, Len(GetFile.FileName) - 3, 1) <> "." Then
    GetFile.FileName = GetFile.FileName & ".bmp"
End If


Select Case Index
        Case 1
            If ViewMode <> 5 Then
             SavePicture view.Image, GetFile.FileName
            Else
             SavePicture DirectX.Image, GetFile.FileName
             MsgBox "Chances are, that didn't work - Sorry!!"
            End If
        Case 2
            RecordScene = True
End Select

CanceledPrintScreen:

End Sub



Private Sub OpView_Click(Index As Integer)
    Select Case Index
        Case 2: CenterView
        Case 3, 4
           If opview(Index).Checked = True Then
               opview(Index).Checked = False
            Else
               opview(Index).Checked = True
            End If
            If opview(4).Checked = True Then
                TopRule.Visible = True
                SideRule.Visible = True
                Corner.Visible = True
            Else
                TopRule.Visible = False
                SideRule.Visible = False
                Corner.Visible = False
            End If
            SortOutScreen
            DrawModel
    End Select
End Sub



Private Sub PresetScale_Change()
    SklMode(2) = True
End Sub

Private Sub PresetScale_GotFocus()
    OpEdit(8).Enabled = False
End Sub

Private Sub PresetScale_LostFocus()
    OpEdit(8).Enabled = True
End Sub

Private Sub QuickSpin_Click(Index As Integer)
     If Index = 0 Then SpinSelected -5
     If Index = 1 Then SpinSelected 5
     If Index = 2 Then SpinSelected -90
     If Index = 3 Then SpinSelected 90
End Sub

Private Sub Sbar_PanelClick(ByVal Panel As MSComctlLib.Panel)
If Panel.Index = 3 And Rotation <> 0 Then
    Rotation = 0
    SBar.Panels(3) = "Cancelled"
    MainTab.Enabled = True
    view.Enabled = True
End If
End Sub

Private Sub Scales_Click()
    PresetScale = Scales.Text
End Sub

Private Sub sclXDim_Click(Index As Integer)
    SklMode(1) = True
End Sub

Private Sub sclXDim_Scroll(Index As Integer)
    SklMode(1) = True
End Sub

Private Sub ShapeList_Click()
    SetNewShapeMenu
    Guide.Visible = False
    If ShapeList.Text = "Roundoid" Or ShapeList.Text = "Torous" Then
        AlineAxis 0, view.ScaleHeight / 2
        Axis.Visible = True
    Else
        Axis.Visible = False
    End If
    DrawGuide
End Sub

Private Sub ShpProp_Click(Index As Integer)
    DrawGuide
End Sub

Private Sub ShpProp_Scroll(Index As Integer)
    DrawGuide
End Sub

Private Sub sldLight_Click(Index As Integer)
    SetLights sldLight(0), sldLight(1)
End Sub

Private Sub sldLight_Scroll(Index As Integer)
    SetLights sldLight(0), sldLight(1)
End Sub

Public Sub StartAnimation()
    TBar3.buttons(4).Value = tbrUnpressed
    TBar3.buttons(5).Value = tbrPressed
    TheSceneOn = 0
    For N = 1 To MnuScene.Count - 1
    If MnuScene(N).Checked = True Then TheSceneOn = N
    Next N
    If TheSceneOn = 0 Then TBar3.buttons(5).Value = tbrPressed: Exit Sub
    frmMain.Label3 = ""
    Key = "Scene" & TheSceneOn
    sceneEdit.Frame.Nodes(Key).Child.FirstSibling.Selected = True
    Num = CountJoints
    Num = Int(Num / 6)
    If Num <> CountJoints Then sceneEdit.SetUpGrid CountJoints
    Animate = True
End Sub

Private Sub Texture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static OldX, OldY
    If Button = 1 Then
        If lstPaintTool.Selected(0) = True Then
            Texture.DrawWidth = LineSize
            Texture.Line (X, Y)-(OldX, OldY), picPallette.Tag
        End If
        If lstPaintTool.Selected(1) = True Then
            Texture.TextWidth "Hi there!"
        End If
    End If
    OldX = X: OldY = Y
End Sub

Private Sub tmaLightEffect_Timer()
    
    If ViewMode < 3 Then Exit Sub
    
    If Rotation <> 0 Then
        SBar.Panels(3) = Rotation & " (click to cancel)"
        Rotation = Rotation - 1
        If frmRecord.Dire(0) = 1 Then Angle1 = Angle1 - (360 / frmRecord.Angle)
        If frmRecord.Dire(1) = 1 Then Angle2 = Angle2 - (360 / frmRecord.Angle)
        If frmRecord.Dire(2) = 1 Then Angle3 = Angle3 - (360 / frmRecord.Angle)
        DrawModel
        SavePicture view.Image, Mid(frmRecord.GetFile.FileName, 1, Len(frmRecord.GetFile.FileName) - 4) & Rotation & ".bmp"
        If Rotation = 0 Then
            SBar.Panels(3) = "Finished"
            MainTab.Enabled = True
            view.Enabled = True
        End If
    End If
    
    If Animate = True And ViewMode > 3 Then
        If sceneEdit.Frame.SelectedItem.Key = sceneEdit.Frame.SelectedItem.LastSibling.Key Then
            sceneEdit.Frame.Nodes(sceneEdit.Frame.SelectedItem.FirstSibling.Key).Selected = True
            GetThis = sceneEdit.Frame.SelectedItem.Key
            sceneEdit.Label3.Caption = "Working file - " & GetThis
            If RecordScene = True Then
                SavePicture view.Image, Mid(GetFile.FileName, 1, Len(GetFile.FileName) - 4) & GetThis & ".bmp"
                Animate = False
                MsgBox "The entire scene has been copied to BMP format"
            End If
            RecordScene = False
         Else
            GetThis = sceneEdit.Frame.SelectedItem.Next.Key
            sceneEdit.Frame.Nodes(GetThis).Selected = True
            sceneEdit.Label3.Caption = "Working file - " & GetThis
            If RecordScene = True Then
                SavePicture view.Image, Mid(GetFile.FileName, 1, Len(GetFile.FileName) - 4) & GetThis & ".bmp"
            End If
        End If
        SBar.Panels(3) = "Working file -" & GetThis
        DrawModel
    End If
    If ViewMode = 5 Then
        lstLights.Tag = lstLights.Tag + 1
        LightStyleOn = lstLights.ListIndex
        If LightStyleOn = 0 Then Exit Sub
        If lstLights.Tag > Len(LightStyle(LightStyleOn)) Then lstLights.Tag = 1
        shade = Asc(Mid(LightStyle(LightStyleOn), Val(lstLights.Tag), 1)) - 65
        shade = (100 / 26) * shade
        SetLights sldLight(0).Value, shade
    
    End If
End Sub

Private Sub TopRule_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    view.AutoRedraw = False
    If (X - hofs < Marker(2, 1) * Zoom + 3) And (X - hofs > Marker(2, 1) * Zoom - 3) Then Marker(2, 3) = 1
    If (X - hofs < Marker(2, 2) * Zoom + 3) And (X - hofs > Marker(2, 2) * Zoom - 3) Then Marker(2, 3) = 0: Marker(2, 4) = 1
End Sub

Private Sub TopRule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then Exit Sub
    If Marker(2, 3) = 1 Then Marker(2, 1) = (X - hofs) / Zoom
    If Marker(2, 4) = 1 Then Marker(2, 2) = (X - hofs) / Zoom
    If Marker(2, 1) < Marker(2, 2) Then
        If Button = 1 Then Marker(2, 1) = Marker(2, 2)
        If Button = 2 Then Marker(2, 2) = Marker(2, 1)
    End If
    SideRule.ToolTipText = Marker(2, 1) - Marker(2, 2)
    view.Cls
    DrawRulers
    view.DrawStyle = 2: view.Line (X + 17, 0)-(X + 17, view.ScaleWidth): view.DrawStyle = 0
End Sub

Private Sub TopRule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Marker(2, 3) = 0: Marker(2, 4) = 0
    DrawRulers
    view.AutoRedraw = True
End Sub

Private Sub sideRule_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    view.AutoRedraw = False
    If (Y - vofs < Marker(1, 1) * Zoom + 3) And (Y - vofs > Marker(1, 1) * Zoom - 3) Then Marker(1, 3) = 1
    If (Y - vofs < Marker(1, 2) * Zoom + 3) And (Y - vofs > Marker(1, 2) * Zoom - 3) Then Marker(1, 3) = 0: Marker(1, 4) = 1
End Sub

Private Sub sideRule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then Exit Sub
    If Marker(1, 3) = 1 Then Marker(1, 1) = (Y - vofs) / Zoom
    If Marker(1, 4) = 1 Then Marker(1, 2) = (Y - vofs) / Zoom
    If Marker(1, 1) < Marker(1, 2) Then
        If Button = 1 Then Marker(1, 1) = Marker(1, 2)
        If Button = 2 Then Marker(1, 2) = Marker(1, 1)
    End If
    SideRule.ToolTipText = Marker(1, 1) - Marker(1, 2)
    view.Cls
    DrawRulers
    view.DrawStyle = 2: view.Line (0, Y + 17)-(view.ScaleWidth, Y + 17): view.DrawStyle = 0
End Sub

Private Sub sideRule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Marker(1, 3) = 0: Marker(1, 4) = 0
    view.AutoRedraw = True
    DrawRulers
End Sub

Private Sub createobject()
        Guide.Visible = False
        view.Cls
        Guide.Tag = "False"
'        SetNewShapeMenu
        If Index = 0 Then
            If ShapeList.Text = "Torous" Then Add_a.Tourous
            If ShapeList.Text = "Cube" Then Add_a.Cube
            If ShapeList.Text = "Plane" Then Add_a.Plane
            If ShapeList.Text = "Prism" Then Add_a.Prism
            If ShapeList.Text = "Pyramid" Then Add_a.Pyramid
            If ShapeList.Text = "Dimond" Then Add_a.Dimond
            If ShapeList.Text = "Sphere" Then Add_a.Sphere
            If ShapeList.Text = "Roundoid" Then Add_a.Roundoid (LineLength)
        End If
        LineLength = 0
        DrawModel
        frmMain.view.DrawStyle = 0
        frmMain.view.AutoRedraw = True
End Sub

Private Sub mnuEditPopUp_click()
 OpEditPopUp(1).Caption = OpEdit(1).Caption
End Sub

Private Sub txtCustomAngle_GotFocus()
    optSpin = True
    OpEdit(8).Enabled = False
End Sub

Private Sub txtCustomAngle_LostFocus()
    OpEdit(8).Enabled = True
End Sub





Private Sub udnChangeAngle_Change()
    optSpin = True
End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xof = (frmMain.view.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.view.ScaleHeight / 2) - frmMain.UDBar
    X = Snaped(X - xof) + xof
    Y = Snaped(Y - yof) + yof
    Guide.Visible = False
    
    If ViewMode = 6 Then
        If Almostt(X, Y, Texture.Left + Texture.Width, Texture.Top + Texture.Height, 10) = True Then
            Guide.x1 = X: Guide.y1 = Y
            Guide.Tag = "DragTex"
        End If
        Exit Sub
    End If
    
    If CurrentSideBar = 4 Then
            RotateMove = 0
            If Button = 1 Then RotateMove = 1
    End If
    
    Select Case Edit_Tool
        Case 1
            GotVert = False
            If Button = 1 Then
                DontDeselect = False
                If chkVert = 0 Then
                    ClickSelect X, Y, Shift
                Else
                    For N = 1 To cstTotalObjects
                        If Object(N).Selected = True And Object(N).Used = True Then
                            For M = 1 To Object(N).VertexCount
                                If ViewMode = 1 Then XX = Object(N).Vertex(M).X * Zoom + xof: YY = Object(N).Vertex(M).Z * Zoom + yof
                                If ViewMode = 2 Then XX = Object(N).Vertex(M).X * Zoom + xof: YY = Object(N).Vertex(M).Y * Zoom + yof
                                If ViewMode = 3 Then XX = Object(N).Vertex(M).Z * Zoom + xof: YY = Object(N).Vertex(M).Y * Zoom + yof
                                If Almostt(Int(X), Int(Y), Int(XX), Int(YY), 7) = True Then
                                    If Object(N).Vertex(M).Selected = True Then
                                        Object(N).Vertex(M).Selected = False
                                    Else
                                        Object(N).Vertex(M).Selected = True
                                    End If
                                    GotVert = True
                                End If
                            Next M
                        End If
                    Next N
                End If
                DontDeselect = False
                If X > Model.Min.X * Zoom + xof - 5 And X < Model.Max.X * Zoom + xof + 5 Then
                    If Y > Model.Min.Y * Zoom + yof - 5 And Y < Model.Max.Y * Zoom + yof + 5 Then
                        DontDeselect = True
                    End If
                End If
                If Shift = 0 And DontDeselect = False Then
                    If chkVert = 0 Then
                        DeselectAll
                    Else
                        If Shift = 0 Then DeselectAllVertecies
                    End If
                End If
                If chkVert = 1 Then DontDeselect = False
                If DontDeselect = True And GotVert = False Then view.MousePointer = 15
            End If
            If Button = 1 Then Guide.x1 = X: Guide.y1 = Y
            
          Case 2
            If Button = 2 And ShapeList.Text <> "Roundoid" And EditLine(0) = False Then
                If Guide.Tag = "True" Then
                    createobject
                Else
                    PopupMenu mnuEditPopUp
                End If
                Exit Sub
            End If
            If ViewMode > 3 Then Exit Sub
            If Almostt(X, Y, Guide.x1, Guide.y1, 3) = True Then MoveLine = 1
            If Almostt(X, Y, Guide.x1, Guide.y2, 3) = True Then MoveLine = 2
            If Almostt(X, Y, Guide.x2, Guide.y2, 3) = True Then MoveLine = 3
            If Almostt(X, Y, Guide.x2, Guide.y1, 3) = True Then MoveLine = 4
            If Guide.Tag = "False" Then
                Guide.x2 = X: Guide.y2 = Y
                Guide.Tag = "True"
                Guide.x1 = X
                Guide.y1 = Y
                MoveLine = 1
                DrawGuide
            End If
            If ShapeList.Text = "Roundoid" Or ShapeList.Text = "Torous" Then
                If EditLine(2).Value = True Then
                    Guide.x1 = X: Guide.y1 = Y
                End If
                If EditLine(1).Value = True And LineLength <> 255 And ShapeList.Text = "Roundoid" Then
                    ShowTime = (100 / 255) * LineLength
                    LineLength = LineLength + 1
                    StoreLine(LineLength).X = X
                    StoreLine(LineLength).Y = Y
                    Guide.x1 = X: Guide.y1 = Y
                    If ShapeList.Text <> "Torous" Then Guide.Visible = True
                    DrawGuide
                End If
                If Button = 2 And EditLine(0).Value = True Then
                    If Axis.Tag = "YX" Then
                        Axis.Tag = "Y"
                    ElseIf Axis.Tag = "Y" Then
                        Axis.Tag = "X"
                    ElseIf Axis.Tag = "XY" Then
                        Axis.Tag = "X"
                    Else
                        Axis.Tag = "Y"
                    End If
                    AlineAxis X, Y: DrawGuide
                End If
            End If
        Case 3
            If Button = 2 Then
                PopupMenu mnuEditPopUp
                Exit Sub
            End If
            '##########################################
            If edMode(3) = True Then
                OpEdit(1).Caption = "Undo delete faces": StoreCurentPosition:        frmMain.OpEdit(1).Enabled = True
                For NowOn = 1 To cstTotalObjects
                    If Object(NowOn).Selected = True Then
                        FaceON = 1
                        For N = 1 To Object(NowOn).FaceCount
                            EdgeCount = Object(NowOn).Face(FaceON)
                            FaceON = FaceON + 1
                            CenX = 0: CenY = 0: CenZ = 0
                            For M = 1 To EdgeCount
                                edge = Object(NowOn).Face(FaceON)
                                FaceON = FaceON + 1
                                CenX = CenX + Object(NowOn).Vertex(edge).X
                                CenY = CenY + Object(NowOn).Vertex(edge).Y
                                CenZ = CenZ + Object(NowOn).Vertex(edge).Z
                            Next M
                            CenX = CenX / EdgeCount
                            CenY = CenY / EdgeCount
                            CenZ = CenZ / EdgeCount
                            MOveThis = False
                            If ViewMode = 1 And Almostt((CenX * Zoom + xof), (CenZ * Zoom + yof), X, Y, 6) = True Then MOveThis = True
                            If ViewMode = 2 And Almostt((CenX * Zoom + xof), (CenY * Zoom + yof), X, Y, 6) = True Then MOveThis = True
                            If ViewMode = 3 And Almostt((CenZ * Zoom + xof), (CenY * Zoom + yof), X, Y, 6) = True Then MOveThis = True
                            If MOveThis = True Then
                                DeleteFace NowOn, N
                                DrawModel
                                Exit Sub
                            End If
                        Next N
                    End If
                Next NowOn
            End If
            '##########################################
            If edMode(0) = True Then
                Scaleline = 0
                If Almostt(X, Y, Model.Max.X * Zoom + xof, Model.Max.Y * Zoom + yof, 14) = True Then Scaleline = 1: Guide.x1 = Model.Min.X + xof: Guide.y1 = Model.Min.Y + yof: Guide.x2 = Model.Max.X + xof: Guide.y2 = Model.Max.Y + yof
                If Almostt(X, Y, Model.Min.X * Zoom + xof, Model.Max.Y * Zoom + yof, 14) = True Then Scaleline = 2: Guide.x1 = Model.Max.X + xof: Guide.y1 = Model.Min.Y + yof: Guide.x2 = Model.Min.X + xof: Guide.y2 = Model.Max.Y + yof
                If Almostt(X, Y, Model.Min.X * Zoom + xof, Model.Min.Y * Zoom + yof, 14) = True Then Scaleline = 3: Guide.x1 = Model.Max.X + xof: Guide.y1 = Model.Max.Y + yof: Guide.x2 = Model.Min.X + xof: Guide.y2 = Model.Min.Y + yof
                If Almostt(X, Y, Model.Max.X * Zoom + xof, Model.Min.Y * Zoom + yof, 14) = True Then Scaleline = 4: Guide.x1 = Model.Min.X + xof: Guide.y1 = Model.Max.Y + yof: Guide.x2 = Model.Max.X + xof: Guide.y2 = Model.Min.Y + yof
                If Almostt(X, Y, (Guide.x1 + Guide.x2) / 2, Guide.y1, 3) = True Then Scaleline = 5
                If Almostt(X, Y, (Guide.x1 + Guide.x2) / 2, Guide.y2, 3) = True Then Scaleline = 6
                If Almostt(X, Y, Guide.x1, (Guide.y1 + Guide.y2) / 2, 3) = True Then Scaleline = 7
                If Almostt(X, Y, Guide.x2, (Guide.y1 + Guide.y2) / 2, 3) = True Then Scaleline = 8
                If Scaleline = 0 Then Exit Sub
            Else
                Guide.x1 = X: Guide.y1 = Y
            End If
            view.AutoRedraw = False
        Case 4
            If Button = 2 Then
                PopupMenu mnuEditPopUp
                Exit Sub
            End If
            Scaleline = 0
            If Almostt(X, Y, Model.Max.X + xof, Model.Max.Y + yof, 5) = True Then Scaleline = 1: Guide.x1 = Model.Min.X + xof: Guide.y1 = Model.Min.Y + yof: Guide.x2 = Model.Max.X + xof: Guide.y2 = Model.Max.Y + yof: view.MousePointer = 8
            If Almostt(X, Y, Model.Min.X + xof, Model.Max.Y + yof, 5) = True Then Scaleline = 2: Guide.x1 = Model.Max.X + xof: Guide.y1 = Model.Min.Y + yof: Guide.x2 = Model.Min.X + xof: Guide.y2 = Model.Max.Y + yof: view.MousePointer = 6
            If Almostt(X, Y, Model.Min.X + xof, Model.Min.Y + yof, 5) = True Then Scaleline = 3: Guide.x1 = Model.Max.X + xof: Guide.y1 = Model.Max.Y + yof: Guide.x2 = Model.Min.X + xof: Guide.y2 = Model.Min.Y + yof: view.MousePointer = 8
            If Almostt(X, Y, Model.Max.X + xof, Model.Min.Y + yof, 5) = True Then Scaleline = 4: Guide.x1 = Model.Min.X + xof: Guide.y1 = Model.Max.Y + yof: Guide.x2 = Model.Max.X + xof: Guide.y2 = Model.Min.Y + yof: view.MousePointer = 6
            
            If Almostt(X, Y, Model.Max.X + xof, (Int(Model.Min.Y + Model.Max.Y) * 0.5) + yof, 6) = True Then Scaleline = 5: Guide.x1 = Model.Min.X + xof: Guide.y1 = Model.Min.Y + yof: Guide.x2 = Model.Max.X + xof: Guide.y2 = Model.Max.Y + yof: view.MousePointer = 9
            If Almostt(X, Y, Model.Min.X + xof, (Int(Model.Min.Y + Model.Max.Y) * 0.5) + yof, 6) = True Then Scaleline = 6: Guide.x1 = Model.Max.X + xof: Guide.y1 = Model.Min.Y + yof: Guide.x2 = Model.Min.X + xof: Guide.y2 = Model.Max.Y + yof: view.MousePointer = 9
            If Almostt(X, Y, (Int(Model.Min.X + Model.Max.X) * 0.5) + xof, Model.Max.Y + yof, 6) = True Then Scaleline = 7: Guide.x1 = Model.Min.X + xof: Guide.y1 = Model.Min.Y + yof: Guide.x2 = Model.Max.X + xof: Guide.y2 = Model.Max.Y + yof: view.MousePointer = 7
            If Almostt(X, Y, (Int(Model.Min.X + Model.Max.X) * 0.5) + xof, Model.Min.Y + yof, 6) = True Then Scaleline = 8: Guide.x1 = Model.Min.X + xof: Guide.y1 = Model.Max.Y + yof: Guide.x2 = Model.Max.X + xof: Guide.y2 = Model.Min.Y + yof: view.MousePointer = 7
            If Scaleline = 0 Then Exit Sub
            view.AutoRedraw = False
        Case 5, 9
            Guide.x1 = X: Guide.y1 = Y
            Guide.x2 = X: Guide.y2 = Y
            Guide.Visible = True
            CustomX = X - xof
            CustomY = Y - yof
            optGetCenter(2).Caption = "Custom " & CustomX & "," & CustomY
        Case 6
            If Button = 2 Then PopupMenu mnuEditPopUp: Exit Sub
            If ViewMode > 3 Then Exit Sub
            If opChangeJ = True Then
                Model.Saved = False
                Guide.x1 = X
                Guide.y1 = Y
                JOn = JointOver(Xstart, Ystart)
                If JOn <> 0 Then
                    Guide.Visible = True
                    Joints.Nodes(BaseFrame(JOn).Key).Selected = True
                End If
            End If
            If opAddJ = True And Button = 1 Then
                Model.Saved = False
                Add = GetJoint
                If Add = 0 Then
                  MsgBox "You have the maximum number of joints allowed"
                  Exit Sub
                End If
                If ViewMode = 1 Then
                   BaseFrame(Add).Position.X = (X - xof) / Zoom
                   BaseFrame(Add).Position.Y = 0
                   BaseFrame(Add).Position.Z = (Y - yof) / Zoom
                End If
                If ViewMode = 2 Then
                   BaseFrame(Add).Position.X = (X - xof) / Zoom
                   BaseFrame(Add).Position.Y = (Y - yof) / Zoom
                   BaseFrame(Add).Position.Z = 0
                End If
                If ViewMode = 3 Then
                   BaseFrame(Add).Position.X = 0
                   BaseFrame(Add).Position.Y = (Y - yof) / Zoom
                   BaseFrame(Add).Position.Z = (X - xof) / Zoom
                End If
                BaseFrame(Add).Name = "Joint " & Add
                BaseFrame(Add).Used = True
                BaseFrame(Add).Key = "Joint " & Add
                BaseFrame(Add).Target = Joints.SelectedItem.Key
                Rel = Joints.SelectedItem.Key
                Joints.Nodes.Add Rel, 4, "Joint " & Add, "Joint " & Add, 4
                Joints.Nodes("Joint " & Add).EnsureVisible
                DeselectAll
                BaseFrame(Add).Selected = True
                Joints.Nodes("Joint " & Add).Selected = True
                Joints.Nodes("Joint " & Add).Tag = 2
                Model.Saved = False
                DrawModel
            End If
    End Select
End Sub

Private Sub View_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xof = (frmMain.view.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.view.ScaleHeight / 2) - frmMain.UDBar
    X = Snaped(X - xof) + xof
    Y = Snaped(Y - yof) + yof
    Xstart = X
    Ystart = Y
    If ViewMode = 4 Then
        Guide.x1 = Guide.x2
        Guide.y1 = Guide.y2
        Guide.x2 = X
        Guide.y2 = Y
        If Button = 1 Then
            Angle1 = Angle1 + (Guide.y1 - Guide.y2)
            Angle2 = Angle2 + (Guide.x1 - Guide.x2)
            Angle1 = Angle1 Mod 360
            Angle2 = Angle2 Mod 360
            Angle3 = Angle3 Mod 360
            If Angle1 < 0 Then Angle1 = Angle1 + 360
            If Angle2 < 0 Then Angle2 = Angle2 + 360
            If Angle3 < 0 Then Angle3 = Angle3 + 360
            DrawModel
            frmMain.view.Refresh
        End If
    End If
    If ViewMode < 4 Then
      Select Case Edit_Tool
            Case 1
                Guide.x2 = X: Guide.y2 = Y
                If DontDeselect = False Then
                    If Button = 1 And Almostt(Guide.x1, Guide.y1, X, Y, 3) = False Then
                        view.AutoRedraw = False
                        view.Cls
                        view.DrawStyle = 2
                        frmMain.view.Line (Guide.x1, Guide.y1)-(Guide.x2, Guide.y2), , B
                    End If
                Else
                    If Button = 1 Then
                        view.AutoRedraw = False
                        view.Cls
                        view.DrawStyle = 2
                        
                        If X < -25 Then LRBar = LRBar - 20: Guide.x1 = Guide.x1 + 20
                        If X > view.ScaleWidth + 25 Then LRBar = LRBar + 20: Guide.x1 = Guide.x1 - 20
                        If Y < -25 Then UDBar = UDBar - 20: Guide.y1 = Guide.y1 + 20
                        If Y > view.ScaleHeight + 25 Then UDBar = UDBar + 20: Guide.y1 = Guide.y1 - 20
                        
                        X = Guide.x1 - Guide.x2
                        Y = Guide.y1 - Guide.y2
                        Drawings.DrawGuideLines X, Y
                    End If
                End If
            Case 2
                If ShapeList.Text = "Torous" Then
                    If EditLine(1) = True Then
                        If MoveLine = 1 Then Guide.x1 = X: Guide.y1 = Y: DrawGuide
                        If MoveLine = 2 Then Guide.x1 = X: Guide.y2 = Y: DrawGuide
                        If MoveLine = 3 Then Guide.x2 = X: Guide.y2 = Y: DrawGuide
                        If MoveLine = 4 Then Guide.x2 = X: Guide.y1 = Y: DrawGuide
                    End If
                    If Button = 1 And EditLine(0).Value = True Then
                        AlineAxis X, Y:  DrawGuide
                    End If
                    Exit Sub
                End If
                If ShapeList.Text <> "Roundoid" Then
                    If MoveLine = 1 Then Guide.x1 = X: Guide.y1 = Y: DrawGuide
                    If MoveLine = 2 Then Guide.x1 = X: Guide.y2 = Y: DrawGuide
                    If MoveLine = 3 Then Guide.x2 = X: Guide.y2 = Y: DrawGuide
                    If MoveLine = 4 Then Guide.x2 = X: Guide.y1 = Y: DrawGuide
                Else
                If EditLine(1).Value = True And Button = 2 And LineLength <> 255 Then
                    ShowTime = (100 / 255) * LineLength
                    LineLength = LineLength + 1
                    StoreLine(LineLength).X = X
                    StoreLine(LineLength).Y = Y
                    Guide.x1 = X: Guide.y1 = Y
                    Guide.Visible = True
                    DrawModel
                    DrawGuide
                End If

                Guide.x2 = X: Guide.y2 = Y
                view.AutoRedraw = True
                If Button = 1 And EditLine(2).Value = True Then
                    
                    For N = 1 To LineLength
                        StoreLine(N).X = StoreLine(N).X + Guide.x2 - Guide.x1
                        StoreLine(N).Y = StoreLine(N).Y + Guide.y2 - Guide.y1
                    Next N
                    Guide.x1 = X: Guide.y1 = Y
 
                    DrawModel
                    DrawGuide
                End If
                If Button = 1 And EditLine(0).Value = True Then
                    DrawModel
                    AlineAxis X, Y:  DrawGuide
                End If
            End If
        Case 3
            If Button <> 0 Then
                If edMode(0) = True And Scaleline = 0 Then Exit Sub
                view.DrawStyle = 2
                view.Cls
                If edMode(0) = True Then
                    Select Case Scaleline
                        Case 1, 3
                            If Scaleline = 1 Or Scaleline = 3 Then
                                view.Line (Model.Max.X + xof, Model.Min.Y + yof)-(X, Y), Colours(3)
                                view.Line (Model.Min.X + xof, Model.Max.Y + yof)-(X, Y), Colours(3)
                            End If
                        Case 2, 4
                            If Scaleline = 2 Or Scaleline = 4 Then
                                view.Line (Model.Min.X + xof, Model.Min.Y + yof)-(X, Y), Colours(3)
                                view.Line (Model.Max.X + xof, Model.Max.Y + yof)-(X, Y), Colours(3)
                            End If
                        Case 99
                                view.Line (Model.Min.X + xof, Model.Max.Y + yof)-(X, Y), Colours(3)
                                view.Line (Model.Max.X + xof, Model.Max.Y + yof)-(X, Y), Colours(3)
                        End Select
                Else
                    Guide.x2 = X: Guide.y2 = Y
                End If
                view.DrawStyle = 0
            End If
        Case 4
            If Scaleline = 0 Then Exit Sub
            If Button = 1 Then
                view.Cls
                view.DrawStyle = 2
                Select Case Scaleline
                    Case 1, 2, 3, 4
                        view.Line (Guide.x1, Guide.y1)-(X, Y), Colours(3), B
                    Case 5, 6
                        view.Line (Guide.x1, Guide.y1)-(X, Guide.y2), Colours(3), B
                    Case 7, 8
                        view.Line (Guide.x1, Guide.y1)-(Guide.x2, Y), Colours(3), B
                End Select
                view.DrawStyle = 0
            End If
        Case 5
            If Button <> 0 Then
                Guide.x2 = X
                Guide.y2 = Y
                txtCustomAngle = Int(GetAngle(Guide.x2 - Guide.x1, Guide.y1 - Guide.y2))
                optSpin = True
            End If
        Case 6
            If opChangeJ = True Then
                Guide.x2 = X
                Guide.y2 = Y
            End If
        End Select
    End If
End Sub

Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xof = (frmMain.view.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.view.ScaleHeight / 2) - frmMain.UDBar
    X = Snaped(X - xof) + xof
    Y = Snaped(Y - yof) + yof
    
    If ViewMode = 6 And Guide.Tag = "DragTex" Then
            Texture.Width = X - Texture.Left
            Texture.Height = Y - Texture.Top
            Guide.Tag = ""
            view.Cls
            view.Circle (Texture.Left + Texture.Width, Texture.Top + Texture.Height), 5
        Exit Sub
    End If
    
    view.MousePointer = 0
    MoveLine = 0
    Select Case Edit_Tool
        Case 1
            If Button = 1 And DontDeselect = True Then
                OpEdit(1).Enabled = True
                OpEdit(1).Caption = "Undo move"
                StoreCurentPosition
                OpEditPopUp(1).Enabled = True
                view.Cls
                view.AutoRedraw = True
                MoveSelected Guide.x2 - Guide.x1, Guide.y2 - Guide.y1
            End If
            If DontDeselect = False Then
                If Button = 1 Then
                    If Almostt(Guide.x1, Guide.y1, Guide.x2, Guide.y2, 3) = True Then
                        If chkVert = 0 Then ClickSelect X, Y, Shift
                    Else
                        If chkVert = 0 Then
                                view.Cls: view.AutoRedraw = True
                                BoxBandSelect Guide.x1, Guide.y1, Guide.x2, Guide.y2, Shift
                        Else
                            view.Cls
                            view.AutoRedraw = True
                            For N = 1 To cstTotalObjects
                                If Object(N).Selected = True Then
                                    For M = 1 To Object(N).VertexCount
                                        If ViewMode = 1 Then XX = Object(N).Vertex(M).X * Zoom + xof: YY = Object(N).Vertex(M).Z * Zoom + yof
                                        If ViewMode = 2 Then XX = Object(N).Vertex(M).X * Zoom + xof: YY = Object(N).Vertex(M).Y * Zoom + yof
                                        If ViewMode = 3 Then XX = Object(N).Vertex(M).Z * Zoom + xof: YY = Object(N).Vertex(M).Y * Zoom + yof
                                        x1 = Guide.x1: y1 = Guide.y1
                                        x2 = Guide.x2: y2 = Guide.y2
                                        If x1 > x2 Then temp = x1: x1 = x2: x2 = temp
                                        If y1 > y2 Then temp = y1: y1 = y2: y2 = temp
                                        If (XX > x1 And XX < x2) And (YY > y1 And YY < y2) Then
                                            If Object(N).Vertex(M).Selected = True Then
                                                If Shift = 2 Then Object(N).Vertex(M).Selected = False
                                            Else
                                                Object(N).Vertex(M).Selected = True
                                            End If
                                        End If
                                    Next M
                                End If
                            Next N
                                
                        End If
                    End If
                End If
            End If
            optJoint.Enabled = False
            If CountSelectedObject > 1 Then optJoint.Enabled = True
            DrawModel
            If Button = 2 Then
                PopupMenu mnuEditPopUp
                Exit Sub
            End If
        Case 2
        Case 3
            If Button = 2 And Almostt(Guide.x1, Guide.y1, Guide.x2, Guide.y2, 3) = True Then
                PopupMenu mnuEdit
                Exit Sub
            End If
            '########################################################
            If edMode(4) = True Or edMode(5) = True Then
            
                If edMode(4) = True Then
                    OpEdit(1).Caption = "Undo bend faces"
                    OpEdit(1).Enabled = True
                    OpEditPopUp(1).Enabled = True
                    StoreCurentPosition
                End If
            
                If edMode(5) = True Then
                    OpEdit(1).Caption = "Undo extend faces"
                    OpEdit(1).Enabled = True
                    OpEditPopUp(1).Enabled = True
                    StoreCurentPosition
                End If
            
            
                For NowOn = 1 To cstTotalObjects
                    If Object(NowOn).Selected = True Then
                        FaceON = 1
                        For N = 1 To Object(NowOn).FaceCount
                            EdgeCount = Object(NowOn).Face(FaceON)
                            FaceON = FaceON + 1
                            CenX = 0: CenY = 0: CenZ = 0
                            For M = 1 To EdgeCount
                                edge = Object(NowOn).Face(FaceON)
                                FaceON = FaceON + 1
                                CenX = CenX + Object(NowOn).Vertex(edge).X
                                CenY = CenY + Object(NowOn).Vertex(edge).Y
                                CenZ = CenZ + Object(NowOn).Vertex(edge).Z
                            Next M
                            CenX = CenX / EdgeCount
                            CenY = CenY / EdgeCount
                            CenZ = CenZ / EdgeCount
                            MOveThis = False
                            If ViewMode = 1 And Almostt((CenX * Zoom + xof), (CenZ * Zoom + yof), Guide.x1, Guide.y1, 6) = True Then MOveThis = True: MidFace = CenY / Zoom - xof
                            If ViewMode = 2 And Almostt((CenX * Zoom + xof), (CenY * Zoom + yof), Guide.x1, Guide.y1, 6) = True Then MOveThis = True: MidFace = CenZ / Zoom - xof
                            If ViewMode = 3 And Almostt((CenZ * Zoom + xof), (CenY * Zoom + yof), Guide.x1, Guide.y1, 6) = True Then MOveThis = True: MidFace = CenX / Zoom - xof
                            If MOveThis = True Then
                                If edMode(4) = True Then BendFace NowOn, N, MidFace
                                If edMode(5) = True Then ExtendFace NowOn, N, MidFace
                                Exit Sub
                            End If
                        Next N
                    End If
                Next NowOn
            End If
            '########################################################
            If edMode(2) = True Then
                OpEdit(1).Caption = "Undo move vertex"
                OpEdit(1).Enabled = True
                OpEditPopUp(1).Enabled = True
                StoreCurentPosition
                For NowOn = 1 To cstTotalObjects
                    If Object(NowOn).Selected = True Then
                        FaceON = 1
                        For N = 1 To Object(NowOn).FaceCount
                            EdgeCount = Object(NowOn).Face(FaceON)
                            FaceON = FaceON + 1
                            CenX = 0: CenY = 0: CenZ = 0
                            For M = 1 To EdgeCount
                                edge = Object(NowOn).Face(FaceON)
                                FaceON = FaceON + 1
                                CenX = CenX + Object(NowOn).Vertex(edge).X
                                CenY = CenY + Object(NowOn).Vertex(edge).Y
                                CenZ = CenZ + Object(NowOn).Vertex(edge).Z
                            Next M
                            CenX = CenX / EdgeCount
                            CenY = CenY / EdgeCount
                            CenZ = CenZ / EdgeCount
                            MOveThis = False
                            If ViewMode = 1 And Almostt((CenX * Zoom + xof), (CenZ * Zoom + yof), Guide.x1, Guide.y1, 6) = True Then MOveThis = True
                            If ViewMode = 2 And Almostt((CenX * Zoom + xof), (CenY * Zoom + yof), Guide.x1, Guide.y1, 6) = True Then MOveThis = True
                            If ViewMode = 3 And Almostt((CenZ * Zoom + xof), (CenY * Zoom + yof), Guide.x1, Guide.y1, 6) = True Then MOveThis = True
                            If MOveThis = True Then
                                 FaceON = FaceON - EdgeCount
                                 If ViewMode = 1 Then Xxx = (Guide.x2 - Guide.x1): Yyy = 0: Zzz = (Guide.y2 - Guide.y1)
                                 If ViewMode = 2 Then Xxx = (Guide.x2 - Guide.x1): Yyy = (Guide.y2 - Guide.y1): Zzz = 0
                                 If ViewMode = 3 Then Yyy = (Guide.y2 - Guide.y1): Zzz = (Guide.x2 - Guide.x1): Xxx = 0
                                 For M = 1 To EdgeCount
                                    edge = Object(NowOn).Face(FaceON)
                                    FaceON = FaceON + 1
                                    Object(NowOn).Vertex(edge).X = Object(NowOn).Vertex(edge).X + Xxx
                                    Object(NowOn).Vertex(edge).Y = Object(NowOn).Vertex(edge).Y + Yyy
                                    Object(NowOn).Vertex(edge).Z = Object(NowOn).Vertex(edge).Z + Zzz
                                Next M
                                If Shift = 0 Then
                                    FindOutLine NowOn
                                    DrawModel
                                    Exit Sub
                                End If
                            End If
                        Next N
                    End If
                Next NowOn
            End If
            
            If edMode(1) = True Then
                For N = 1 To cstTotalObjects
                    If Object(N).Used = True And Object(N).Selected = True Then
                        For M = 1 To Object(N).VertexCount
                            If ViewMode = 1 Then
                                If Almostt(Guide.x1, Guide.y1, (Object(N).Vertex(M).X * Zoom) + xof, (Object(N).Vertex(M).Z * Zoom) + yof, 5) = True Then
                                    Object(N).Vertex(M).X = Object(N).Vertex(M).X + (Guide.x2 - Guide.x1) / Zoom
                                    Object(N).Vertex(M).Z = Object(N).Vertex(M).Z + (Guide.y2 - Guide.y1) / Zoom
                                End If
                            End If
                            If ViewMode = 2 Then
                                If Almostt(Guide.x1, Guide.y1, (Object(N).Vertex(M).X * Zoom) + xof, (Object(N).Vertex(M).Y * Zoom) + yof, 5) = True Then
                                    Object(N).Vertex(M).X = Object(N).Vertex(M).X + (Guide.x2 - Guide.x1) / Zoom
                                    Object(N).Vertex(M).Y = Object(N).Vertex(M).Y + (Guide.y2 - Guide.y1) / Zoom
                                End If
                            End If
                            If ViewMode = 3 Then
                                If Almostt(Guide.x1, Guide.y1, (Object(N).Vertex(M).Z * Zoom) + xof, (Object(N).Vertex(M).Y * Zoom) + yof, 5) = True Then
                                    Object(N).Vertex(M).Z = Object(N).Vertex(M).Z + (Guide.x2 - Guide.x1) / Zoom
                                    Object(N).Vertex(M).Y = Object(N).Vertex(M).Y + (Guide.y2 - Guide.y1) / Zoom
                                End If
                            End If
                        Next M
                    End If
                    FindOutLine (N)
                Next N
            End If
            DrawModel
        Case 4
            If Button = 2 And Almostt(Guide.x1, Guide.y1, Guide.x2, Guide.y2, 3) = True Then
                PopupMenu mnuEditPopUp
                Exit Sub
            End If
            If Scaleline = 0 Then Exit Sub
            view.Cls
            view.AutoRedraw = True
            Select Case Scaleline
                Case 1, 2, 3, 4
                    owidx = (Guide.x1 - Guide.x2)
                    owidy = (Guide.y1 - Guide.y2)
                    nWidx = (Guide.x1 - X)
                    nWidy = (Guide.y1 - Y)
                    If owidx = 0 Or owidy = 0 Then Exit Sub
                    nWidx = nWidx / owidx
                    nWidy = nWidy / owidy
                Case 5, 6
                    owidx = (Guide.x1 - Guide.x2)
                    owidy = (Guide.y1 - Guide.y2)
                    nWidx = (Guide.x1 - X)
                    nWidy = (Guide.y1 - Guide.y2)
                    If owidx = 0 Or owidy = 0 Then Exit Sub
                    nWidx = nWidx / owidx
                    nWidy = nWidy / owidy
                Case 7, 8
                    owidx = (Guide.x1 - Guide.x2)
                    owidy = (Guide.y1 - Guide.y2)
                    nWidx = (Guide.x1 - Guide.x2)
                    nWidy = (Guide.y1 - Y)
                    If owidx = 0 Or owidy = 0 Then Exit Sub
                    nWidx = nWidx / owidx
                    nWidy = nWidy / owidy
            End Select

            
            For NowOn = 1 To cstTotalObjects
                If Object(NowOn).Used = True And Object(NowOn).Selected = True Then
                    If nWidx > 0 And nWidy < 0 Then FlipFaces NowOn
                    If nWidx < 0 And nWidy > 0 Then FlipFaces NowOn
                    For N = 1 To Object(NowOn).VertexCount
                        If frmMain.chkVert = 0 Or Object(NowOn).Vertex(N).Selected = True Then
                            If ViewMode = 1 Then
                                Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X - Guide.x1 + xof
                                Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z - Guide.y1 + yof
                                Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X * nWidx
                                Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z * nWidy
                                Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X + Guide.x1 - xof
                                Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z + Guide.y1 - yof
                            End If
                            If ViewMode = 2 Then
                                Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X - Guide.x1 + xof
                                Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y - Guide.y1 + yof
                                Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X * nWidx
                                Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y * nWidy
                                Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X + Guide.x1 - xof
                                Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y + Guide.y1 - yof
                            End If
                            If ViewMode = 3 Then
                                Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z - Guide.x1 + xof
                                Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y - Guide.y1 + yof
                                Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z * nWidx
                                Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y * nWidy
                                Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z + Guide.x1 - xof
                                Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y + Guide.y1 - yof
                            End If
                        End If
                    Next N
                    FindOutLine (NowOn)
                End If
            Next NowOn
            
            For NowOn = 1 To cstTotalJoints
                If BaseFrame(NowOn).Selected = True Then
                    If ViewMode = 1 Then
                        BaseFrame(NowOn).Position.X = BaseFrame(NowOn).Position.X - Guide.x1 + xof
                        BaseFrame(NowOn).Position.Z = BaseFrame(NowOn).Position.Z - Guide.y1 + yof
                        BaseFrame(NowOn).Position.X = BaseFrame(NowOn).Position.X * nWidx
                        BaseFrame(NowOn).Position.Z = BaseFrame(NowOn).Position.Z * nWidy
                        BaseFrame(NowOn).Position.X = BaseFrame(NowOn).Position.X + Guide.x1 - xof
                        BaseFrame(NowOn).Position.Z = BaseFrame(NowOn).Position.Z + Guide.y1 - yof
                    End If
                    If ViewMode = 2 Then
                        BaseFrame(NowOn).Position.X = BaseFrame(NowOn).Position.X - Guide.x1 + xof
                        BaseFrame(NowOn).Position.Y = BaseFrame(NowOn).Position.Y - Guide.y1 + yof
                        BaseFrame(NowOn).Position.X = BaseFrame(NowOn).Position.X * nWidx
                        BaseFrame(NowOn).Position.Y = BaseFrame(NowOn).Position.Y * nWidy
                        BaseFrame(NowOn).Position.X = BaseFrame(NowOn).Position.X + Guide.x1 - xof
                        BaseFrame(NowOn).Position.Y = BaseFrame(NowOn).Position.Y + Guide.y1 - yof
                    End If
                    If ViewMode = 3 Then
                        BaseFrame(NowOn).Position.Z = BaseFrame(NowOn).Position.Z - Guide.x1 + xof
                        BaseFrame(NowOn).Position.Y = BaseFrame(NowOn).Position.Y - Guide.y1 + yof
                        BaseFrame(NowOn).Position.Z = BaseFrame(NowOn).Position.Z * nWidx
                        BaseFrame(NowOn).Position.Y = BaseFrame(NowOn).Position.Y * nWidy
                        BaseFrame(NowOn).Position.Z = BaseFrame(NowOn).Position.Z + Guide.x1 - xof
                        BaseFrame(NowOn).Position.Y = BaseFrame(NowOn).Position.Y + Guide.y1 - yof
                    End If
                End If
            Next NowOn
            
            
            
            
            DrawModel
        Case 5
            If Button = 2 And Almostt(Guide.x1, Guide.y1, Guide.x2, Guide.y2, 3) = True Then
                PopupMenu mnuEditPopUp
                Exit Sub
            End If
            Guide.Visible = False
            optGetCenter(2) = True
            If Button = 2 Then SpinSelected txtCustomAngle
        Case 6
            If opChangeJ = True Then
                Guide.x2 = X
                Guide.y2 = Y
                Guide.Visible = False
                j1 = JointOver(Guide.x1, Guide.y1)
                j2 = JointOver(Guide.x2, Guide.y2)
                If j2 = j1 Then Exit Sub
                If j1 = 0 Then Exit Sub
                If j2 <> 0 Then
                    FP1 = Joints.Nodes(BaseFrame(j1).Key).FullPath
                    FP2 = Mid(Joints.Nodes(BaseFrame(j2).Key).FullPath, 1, Len(Joints.Nodes(BaseFrame(j1).Key).FullPath))
                End If
                If FP1 = FP2 And j2 <> 0 Then
                    MsgBox "You cannot link a joint to another joint that is directly or indirectly linked to the first joint", 48, "Invalid opperation"
                    Exit Sub
                End If
                
                
                If j2 = 0 Then
                    BaseFrame(j1).Target = "Model"
                Else
                    BaseFrame(j1).Target = BaseFrame(j2).Key
                End If
                Dim JJ1 As Node
                Dim JJ2 As Node
                If j1 = 0 Then Exit Sub
                Set JJ1 = Joints.Nodes(BaseFrame(j1).Key)
                If j2 = 0 Then
                    Set JJ2 = Joints.Nodes("Model")
                Else
                    Set JJ2 = Joints.Nodes(BaseFrame(j2).Key)
                End If
                DrawModel
                DragAndDrop Joints, JJ1, JJ2
                Joints.Nodes(BaseFrame(j1).Key).EnsureVisible
            End If
    End Select

    
End Sub

Private Sub DirectX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MD3D = Button
    m_LastX = X
    m_LastY = Y
End Sub
Public Sub DirectX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MD3D <> 0 Then
        RotateTrackBall CInt(X), CInt(Y)
        RenderScene
        DoEvents
    End If
End Sub
Private Sub DirectX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MD3D = 0
    Speed = 0
End Sub



