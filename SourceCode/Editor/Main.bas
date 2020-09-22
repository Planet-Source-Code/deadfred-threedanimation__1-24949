Attribute VB_Name = "Main"
Option Explicit

'########################################## _
    These set the number of boxes, joints and scenes _
    you can have. The more boexs & joints you have, the _
    slower it goes. The number of scenes has no effect _
    on speed.

    Public Const cstTotalObjects = 200
    Public Const cstTotalVertecies = 8000
    Public Const cstTotalJoints = 100
    Public Const cstTotalScenes = 10
    Public Const cstHandleSize = 3
    
    Public Const GridSize = 4
    Public Const Pye = (22 / 7) * 18.2

    Public Const FaceCenterColour = &HC0&
    Public Const VertexColour = &HC0&
    Public Const ImportTypes$ = "*.asc; *.dat; *.map"

'##########################################

Type SceneDis
    Used As Boolean
    Name As String
    Mode As Byte
    Key As String
End Type

Type ColourDis
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Public Type Scene
    Name As String
    Frames As Integer
    Used As Boolean
End Type

Type Coord
    X As Single
    Y As Single
    Z As Single
    Target As Integer
End Type

Type StoredLine
    X As Integer
    Y As Integer
    Z As Integer
    Angle As Integer
    Height As Integer
End Type

Type VertexDis
    X As Single
    Y As Single
    Z As Single
    Selected As Boolean
    Target As Integer
    TargetName As String
End Type

Type Meshdis
    X As String
    Y As String
    Edges As Byte
    Axis As Byte
    Res As Byte
End Type

Public Type JointDis
    Used As Boolean
    Position As Coord
    NewPosit As Coord
    Target As String
    TargetNum As Integer
    Colour As Double
    Selected As Boolean
    Name As String
    Key As String
    Angle As Coord
    Scale As Coord
    JointType As String
    
    WeaponName As String
    WeaponType As String
    IsAWeapon As Boolean
    HAngle As Integer
    Vangle As Integer
    
End Type

Public Type ModelDis
    Saved As Boolean
    ProjectFileName  As String
    ProjectName As String
    Max As Coord
    Min As Coord
    ProjectCreated As Date
    
    StartShowSideBar As Boolean
    StartViewMode As Byte
    StartAnimated As Boolean
    StartSceneName As Byte
    ShowNotesAtStart As Boolean
    ModelNotes As String
    
    BotForce As Boolean
    BotID As String
    BotWeight As Integer
    BotCost As Integer
    BotDis As String
End Type

Type ObjectDis
    Used As Boolean
    Selected As Boolean
    MeshUsed As Boolean
    VertexCount As Integer
    FaceCount As Integer
    EdgeCount As Integer
    GroupCount As Integer
    Face() As Integer
    Group() As Integer
    Vertex() As VertexDis
    Colour As Long
    Max As Coord
    Min As Coord
End Type


Public Colours(20) As Long

   

Public Animate As Boolean

Public KeepRight(cstTotalJoints, 2) As Integer

Public Object(cstTotalObjects) As ObjectDis
Public BaseFrame(cstTotalJoints) As JointDis
Public SceneSave(cstTotalScenes) As Byte
Public Scenes(cstTotalScenes) As SceneDis
Public SelectMode As Byte
Public Model As ModelDis
Public Rotates As Coord
Public hofs As Integer, vofs As Integer, Zoom As Single
Public Marker(2, 4) As Integer
Public HaHa(cstTotalJoints, 2) As Integer
Public Const Zeye = 800
Public CenterFace() As Integer
Public ViewMode As Byte
Public VertArray() As D3DVECTOR
Public SideFaces() As Long

Public RememberW() As Integer

Declare Function SetWindowPos Lib "User32" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal Cx%, ByVal Cy%, ByVal F%) As Integer

Sub AlwaysOnTop(frmID As Form, OnTop As Integer)
' Pass any non-zero value to Place on top
' Pass zero to remove top-mostness

    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    If OnTop Then
        OnTop = SetWindowPos(frmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(frmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub


Public Sub NewShapeList()

Open App.Path + "\data\Shapes.txt" For Output As #1
Print #1, "Cube"
Print #1, "    {"
Print #1, "        Message , ""A Cube or 3D rectangle"""
Print #1, "    }"
Print #1, "Plane"
Print #1, "    {"
Print #1, "        Message , ""A 2D surface that can be used to make complex 3D shapes"""
Print #1, "        Edges, Any"
Print #1, "        Ofset, Any"
Print #1, "    }"
Print #1, "Prism"
Print #1, "    {"
Print #1, "        Message , ""A 3D prism with a set number of equal sides"""
Print #1, "        Edges, Any"
Print #1, "        Ofset, Any"
Print #1, "        TopFace, Any"
Print #1, "        BottomFace, Any"
Print #1, "    }"
Print #1, "Pyramid"
Print #1, "    {"
Print #1, "        Message , ""A 3D prism"""
Print #1, "        Edges, Any"
Print #1, "        Ofset, Any"
Print #1, "    }"
Print #1, "Dimond"
Print #1, "    {"
Print #1, "        Message , ""Like 2 pyramids with the bottom faces touching"""
Print #1, "        Edges, Any"
Print #1, "        Ofset, Any"
Print #1, "    }"
Print #1, "Sphere"
Print #1, "    {"
Print #1, "        Message , ""A ball"""
Print #1, "        Edges, Any"
Print #1, "        Vertical, Any"
Print #1, "    }"
Print #1, "Torous"
Print #1, "    {"
Print #1, "        Message , ""A dougnut Shape thing wraped around the Axis"""
Print #1, "        Edges, Any"
Print #1, "        Ofset, Any"
Print #1, "        Axis, Any"
Print #1, "        MoveLine , ""Move Axis"""
Print #1, "        ExtendLine , ""Scale profile"""
Print #1, "        Vertical, Any"
Print #1, "    }"
Print #1, "Roundoid"
Print #1, "    {"
Print #1, "        Message , ""A complex spherical Object fitting a profile of you design wraped around the Axis """
Print #1, "        Edges, Any"
Print #1, "        MoveLine , ""MoveAxis"""
Print #1, "        ExtendLine , ""Extend Line"""
Print #1, "        MoveProfile , ""Move profile"""
Print #1, "        Axis, Any"
Print #1, "    }"
Close
End Sub



Public Sub SetNewShapeMenu()
     Dim temp, N, Control, Value
     If frmMain.ShapeList.ListCount = 0 Then
      Open App.Path & "\data\shapes.txt" For Input As #1
       Do Until EOF(1) <> False
        Input #1, temp
        If temp = "{" Then
         Do Until temp = "}"
         Input #1, temp
         Loop
        Else
         frmMain.ShapeList.AddItem temp
        End If
       Loop
      Close
     Else
      For N = 1 To frmMain.ShpProp.Count
       frmMain.ShpName(N).Enabled = False
       frmMain.ShpProp(N).Enabled = False
      Next N: frmMain.Axis.Visible = False
      frmMain.EditLine(0).Enabled = False: frmMain.EditLine(1).Enabled = False: frmMain.EditLine(2).Enabled = False
      Open App.Path & "\data\shapes.txt" For Input As #1
       Do Until EOF(1) <> False
        Input #1, temp
        If temp = frmMain.ShapeList.Text Then
        Input #1, temp
         Do Until Control = "}"
          Input #1, Control
          If Control <> "}" Then Input #1, Value
          If Control = "Edges" Then frmMain.ShpProp(1).Enabled = True: frmMain.ShpName(1).Enabled = True: If Value <> "Any" Then frmMain.ShpProp(1) = Value
          If Control = "Ofset" Then frmMain.ShpProp(2).Enabled = True: frmMain.ShpName(2).Enabled = True: If Value <> "Any" Then frmMain.ShpProp(2) = Value
          If Control = "Total" Then frmMain.ShpProp(3).Enabled = True: frmMain.ShpName(3).Enabled = True: If Value <> "Any" Then frmMain.ShpProp(3) = Value
          If Control = "MoveLine" Then frmMain.EditLine(0).Enabled = True: frmMain.EditLine(0).Caption = Value
          If Control = "ExtendLine" Then frmMain.EditLine(1).Enabled = True: frmMain.EditLine(1).Caption = Value
          If Control = "Axis" Then frmMain.Axis.Visible = True
          If Control = "MoveProfile" Then frmMain.EditLine(2).Enabled = True: frmMain.EditLine(2).Caption = Value
          If Control = "TopFace" Then frmMain.ShpProp(3).Enabled = True: frmMain.ShpName(3).Enabled = True: If Value <> "Any" Then frmMain.ShpProp(3) = Value
          If Control = "BottomFace" Then frmMain.ShpProp(4).Enabled = True: frmMain.ShpName(4).Enabled = True: If Value <> "Any" Then frmMain.ShpProp(4) = Value
          If Control = "Vertical" Then frmMain.ShpProp(5).Enabled = True: frmMain.ShpName(5).Enabled = True: If Value <> "Any" Then frmMain.ShpProp(5) = Value
          If Control = "Message" Then frmMain.SBar.Panels(3).Text = Value
         Loop
        Else
        End If
       Loop
      Close
    End If
End Sub


