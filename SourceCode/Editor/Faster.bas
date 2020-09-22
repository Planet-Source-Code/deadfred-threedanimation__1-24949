Attribute VB_Name = "Faster"
Option Explicit

Dim TempMorph() As Morph
Public Xf As Integer, Yf As Integer

Public BigSkel As Integer
Public Comand As String
Public WorldCount
Public NewFrame As Boolean

Type ColourDis
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Type SkelitonDis
    Origin As Coord
    Target As Integer
    Name As String
    Huuh As String
End Type

Type WeaponDis
    Name As String
    Type As String
    Joint As Integer
    Vangle As Integer
    HAngle As Integer
End Type

Type Objects
    ID As String
    VertexCount As Integer
    Vertex() As Coord
    FaceCount As Integer
    Face() As Integer
    EdgeCount() As Integer
    SkelitonCount As Integer
    Skeliton() As SkelitonDis
    WeaponCount As Byte
    Weapon() As WeaponDis
    ForceSkeliton As Boolean
End Type

Type Morph
    Origin As Coord
    Angle As Coord
    Scale As Coord
End Type

Type WorldDis
    Morph() As Morph
    Node() As Coord
    Angle As Coord
End Type

Public World As WorldDis
Public Comp As Objects
Public TempVert(500) As Coord

Public Function LoadCompressedModel(FileName As String) As Boolean
        On Error GoTo NotAComp
        Dim Test, N As Integer, FaceEdge As Integer, M As Integer, X As Integer
        Dim ForceSkel As Integer, Skel As Integer, Y As Integer, Z As Integer, WeKnow As Integer
        Dim Gun As Integer
        Open FileName For Input As #1
            Input #1, Test, Comp.ID
            Line Input #1, Test
            Line Input #1, Test
            Line Input #1, Test
            Input #1, Comp.VertexCount
            ReDim Comp.Vertex(Comp.VertexCount) As Coord
            Input #1, Comp.FaceCount
            ReDim Comp.Face(Comp.FaceCount, 14) As Integer
            ReDim Comp.EdgeCount(Comp.FaceCount) As Integer
            For N = 1 To Comp.VertexCount
                Input #1, Comp.Vertex(N).X
                Input #1, Comp.Vertex(N).Y
                Input #1, Comp.Vertex(N).Z
                Input #1, Comp.Vertex(N).Target
            Next N
            For N = 1 To Comp.FaceCount
                Input #1, FaceEdge
                Comp.EdgeCount(N) = FaceEdge
                For M = 1 To FaceEdge
                    Input #1, X: Comp.Face(N, M) = X + 1
                Next M
            Next N
            Input #1, ForceSkel
            Input #1, Skel
            If ForceSkel = 1 Then Comp.ForceSkeliton = True
            If Skel > BigSkel Then BigSkel = Skel
            Comp.SkelitonCount = Skel
            ReDim Comp.Skeliton(Skel) As SkelitonDis
            ReDim World.Morph(Skel) As Morph
            For N = 1 To Skel
                Input #1, X
                Input #1, Y
                Input #1, Z
                Input #1, WeKnow
                Comp.Skeliton(WeKnow).Origin.X = X
                Comp.Skeliton(WeKnow).Origin.Y = Y
                Comp.Skeliton(WeKnow).Origin.Z = Z
                Input #1, Comp.Skeliton(WeKnow).Target
                Input #1, Comp.Skeliton(WeKnow).Name
                Input #1, Comp.Skeliton(WeKnow).Huuh
            Next N
            Input #1, Gun
            Comp.WeaponCount = Gun
            ReDim Comp.Weapon(Gun) As WeaponDis
            For N = 1 To Gun
                Input #1, Comp.Weapon(N).Joint
                Input #1, Comp.Weapon(N).Name
                Input #1, Comp.Weapon(N).Type
                Input #1, Comp.Weapon(N).Vangle
                Input #1, Comp.Weapon(N).HAngle
            Next N
        Close
        LoadCompressedModel = True
        Exit Function
NotAComp:
End Function


Public Sub RunEngine(Window As PictureBox)
    Dim Unit As Integer, N As Byte
    ReDim TempMorph(Comp.VertexCount) As Morph
    Rotate
    Window.Cls
    DrawObject Window
End Sub

Private Sub Rotate()
    Dim Oangle1 As Integer, OAngle2 As Integer, Oangle3 As Integer, nn As Integer
    Dim Target As Byte
    Dim Mangle1 As Integer, Mangle2 As Integer, Mangle3 As Integer
    Dim X As Integer, Y As Integer, Z As Integer, XRotated As Integer, YRotated As Integer, ZRotated As Integer
    Oangle1 = World.Angle.X Mod 360
    OAngle2 = World.Angle.Y Mod 360
    Oangle3 = World.Angle.Z Mod 360
    MorphSkeliton
    For nn = 1 To Comp.VertexCount
        Target = Comp.Vertex(nn).Target
        Mangle1 = TempMorph(Target).Angle.X Mod 360
        Mangle2 = TempMorph(Target).Angle.Y Mod 360
        Mangle3 = TempMorph(Target).Angle.Z Mod 360
        X = Comp.Vertex(nn).X - Comp.Skeliton(Target).Origin.X
        Y = Comp.Vertex(nn).Y - Comp.Skeliton(Target).Origin.Y
        Z = Comp.Vertex(nn).Z - Comp.Skeliton(Target).Origin.Z
        XRotated = (COSine(Mangle3) * X - SINe(Mangle3) * Y)
        YRotated = (SINe(Mangle3) * X + COSine(Mangle3) * Y)
        ZRotated = Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = X
        YRotated = COSine(Mangle1) * Y - SINe(Mangle1) * Z
        ZRotated = SINe(Mangle1) * Y + COSine(Mangle1) * Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = (COSine(Mangle2) * X - SINe(Mangle2) * Z) + TempMorph(Target).Origin.X
        YRotated = Y + TempMorph(Target).Origin.Y
        ZRotated = (SINe(Mangle2) * X + COSine(Mangle2) * Z) + TempMorph(Target).Origin.Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = COSine(Oangle3) * X - SINe(Oangle3) * Y
        YRotated = SINe(Oangle3) * X + COSine(Oangle3) * Y
        ZRotated = Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = X
        YRotated = COSine(Oangle1) * Y - SINe(Oangle1) * Z
        ZRotated = SINe(Oangle1) * Y + COSine(Oangle1) * Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = COSine(OAngle2) * X - SINe(OAngle2) * Z
        YRotated = Y
        ZRotated = SINe(OAngle2) * X + COSine(OAngle2) * Z
        X = XRotated: Y = YRotated: Z = ZRotated
        TempVert(nn).X = XRotated
        TempVert(nn).Y = YRotated
        TempVert(nn).Z = ZRotated
    Next nn
End Sub

Private Sub MorphSkeliton()
    Dim N As Integer, Target As Byte, Cx As Integer, Cy As Integer, Cz As Integer, sAngle1 As Integer, sAngle2 As Integer, sAngle3 As Integer
    Dim X As Integer, Y As Integer, Z As Integer
    For N = 1 To Comp.SkelitonCount
        Target = Comp.Skeliton(N).Target
        If Target <> 0 Then
            TempMorph(N).Origin.X = Comp.Skeliton(N).Origin.X
            TempMorph(N).Origin.Y = Comp.Skeliton(N).Origin.Y
            TempMorph(N).Origin.Z = Comp.Skeliton(N).Origin.Z
            Cx = TempMorph(Target).Origin.X
            Cy = TempMorph(Target).Origin.Y
            Cz = TempMorph(Target).Origin.Z
            sAngle1 = TempMorph(Target).Angle.X
            sAngle2 = TempMorph(Target).Angle.Y
            sAngle3 = TempMorph(Target).Angle.Z
            X = TempMorph(Target).Origin.X + Comp.Skeliton(N).Origin.X - Comp.Skeliton(Target).Origin.X + World.Morph(N).Origin.X
            Y = TempMorph(Target).Origin.Y + Comp.Skeliton(N).Origin.Y - Comp.Skeliton(Target).Origin.Y + World.Morph(N).Origin.Y
            Z = TempMorph(Target).Origin.Z + Comp.Skeliton(N).Origin.Z - Comp.Skeliton(Target).Origin.Z + World.Morph(N).Origin.Z
            TempMorph(N).Angle.X = World.Morph(N).Angle.X + sAngle1
            TempMorph(N).Angle.Y = World.Morph(N).Angle.Y + sAngle2
            TempMorph(N).Angle.Z = World.Morph(N).Angle.Z + sAngle3
            sRotate sAngle1, sAngle2, sAngle3, X, Y, Z, Cx, Cy, Cz
            TempMorph(N).Origin.X = Rotates.X
            TempMorph(N).Origin.Y = Rotates.Y
            TempMorph(N).Origin.Z = Rotates.Z
        End If
    Next N
End Sub

Private Sub sRotate(tAngle1, tAngle2, tAngle3, X, Y, Z, Cx, Cy, Cz)
    Dim XRotated As Integer, YRotated As Integer, ZRotated As Integer
    tAngle1 = (tAngle1 Mod 360):    tAngle2 = (tAngle2 Mod 360):    tAngle3 = (tAngle3 Mod 360)
    If tAngle1 < 0 Then tAngle1 = tAngle1 + 360
    If tAngle2 < 0 Then tAngle2 = tAngle2 + 360
    If tAngle3 < 0 Then tAngle3 = tAngle3 + 360
    X = X - Cx:    Y = Y - Cy:    Z = Z - Cz
    XRotated = COSine(tAngle3) * X - SINe(tAngle3) * Y
    YRotated = SINe(tAngle3) * X + COSine(tAngle3) * Y:    ZRotated = Z
    X = XRotated:    Y = YRotated:    Z = ZRotated
    YRotated = COSine(tAngle1) * Y - SINe(tAngle1) * Z
    ZRotated = SINe(tAngle1) * Y + COSine(tAngle1) * Z: XRotated = X
    X = XRotated:    Y = YRotated:    Z = ZRotated:    XRotated = X
    XRotated = COSine(tAngle2) * X - SINe(tAngle2) * Z:    YRotated = Y
    ZRotated = SINe(tAngle2) * X + COSine(tAngle2) * Z
    Rotates.X = XRotated + Cx:    Rotates.Y = YRotated + Cy:    Rotates.Z = ZRotated + Cz
End Sub

Private Sub DrawObject(Window As PictureBox)
    Dim Ner(12, 2) As Integer, N As Integer, M As Integer, s1 As Integer, x1 As Integer, y1 As Integer, z1 As Integer
    Dim xx1 As Integer, yy1 As Integer, xx2 As Integer, yy2 As Integer, X2 As Integer, Y2 As Integer, Z2 As Integer
    On Error GoTo Yikes
    For N = 1 To Comp.FaceCount
        For M = 1 To Comp.EdgeCount(N)
            s1 = Comp.Face(N, M)
            x1 = TempVert(s1).X:  y1 = TempVert(s1).Y: z1 = TempVert(s1).Z
            Ner(M, 1) = Xf + Int(x1 * (Zeye / (Zeye - z1)))
            Ner(M, 2) = Yf + Int(y1 * (Zeye / (Zeye - z1)))
        Next M
        If ((Ner(1, 2) - Ner(3, 2)) * (Ner(2, 1) - Ner(1, 1)) - (Ner(1, 1) - Ner(3, 1)) * (Ner(2, 2) - Ner(1, 2))) < 0 Then
            For M = 1 To Comp.EdgeCount(N) - 1
                xx1 = Ner(M, 1): yy1 = Ner(M, 2)
                xx2 = Ner(M + 1, 1): yy2 = Ner(M + 1, 2)
                Window.Line (xx1, yy1)-(xx2, yy2)
            Next M
            xx1 = Ner(M, 1): yy1 = Ner(M, 2)
            xx2 = Ner(1, 1): yy2 = Ner(1, 2)
            Window.Line (xx1, yy1)-(xx2, yy2)
        End If
    Next N
Yikes:
End Sub
