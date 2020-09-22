Attribute VB_Name = "DirectXh"
'Option Explicit
'declare objects
Dim g_dx As New DirectX7
Dim m_dd As DirectDraw7
Dim m_ddClipper As DirectDrawClipper
Dim m_rm As Direct3DRM3
'declare devices
Dim m_rmDevice As Direct3DRMDevice3
'declare viewports
Dim m_rmViewport As Direct3DRMViewport2
'declare frames
Dim m_rootFrame As Direct3DRMFrame3
Dim m_lightFrame As Direct3DRMFrame3
Dim m_cameraFrame As Direct3DRMFrame3
Dim m_objectFrame As Direct3DRMFrame3
'declare meshes
Dim m_meshBuilder As Direct3DRMMeshBuilder3
Dim Mesh(cstTotalObjects) As Direct3DRMMeshBuilder3
'delare lights
Dim m_light As Direct3DRMLight
Dim m_ambientLight As Direct3DRMLight
'viewport sizes
Dim m_width As Long
Dim m_height As Long
'mousedown positions for rotation
Public m_LastX As Integer
Public m_LastY As Integer
Dim VertArray() As D3DVECTOR 'vertices for the object
Dim SideFaces() As Long     'vertices making the sides of each face


Sub SetLights(Ambiant, Spot)
    m_light.SetColorRGB Spot / 100, Spot / 100, Spot / 100
    m_ambientLight.SetColorRGB Ambiant / 100, Ambiant / 100, Ambiant / 100
    RenderScene
End Sub

Sub InitRM(PicBox As PictureBox)
    Set m_dd = g_dx.DirectDrawCreate("")
    Set m_ddClipper = m_dd.CreateClipper(0)
    m_ddClipper.SetHWnd PicBox.hWnd
    m_width = PicBox.ScaleWidth
    m_height = PicBox.ScaleHeight
    Set m_rm = g_dx.Direct3DRMCreate()
    Set m_rmDevice = m_rm.CreateDeviceFromClipper(m_ddClipper, "", m_width, m_height)
    X = frmMain.SLDQuality.Value
    If X = 1 Then m_rmDevice.SetQuality D3DRMFILL_POINTS
    If X = 2 Then m_rmDevice.SetQuality D3DRMRENDER_WIREFRAME
    If X = 3 Then m_rmDevice.SetQuality D3DRMRENDER_FLAT
    If X = 4 Then m_rmDevice.SetQuality D3DRMRENDER_PHONG
End Sub

Sub SetMode(X)
 If X = 1 Then m_rmDevice.SetQuality D3DRMFILL_POINTS
 If X = 2 Then m_rmDevice.SetQuality D3DRMRENDER_WIREFRAME
 If X = 3 Then m_rmDevice.SetQuality D3DRMRENDER_FLAT
 If X = 4 Then m_rmDevice.SetQuality D3DRMRENDER_PHONG
End Sub

Sub CleanUp()
    Set m_light = Nothing
    Set m_ambientLight = Nothing
    Set m_meshBuilder = Nothing
    Set m_rmViewport = Nothing
    Set m_lightFrame = Nothing
    Set m_cameraFrame = Nothing
    Set m_objectFrame = Nothing
    Set m_rootFrame = Nothing
    Set m_rmDevice = Nothing
    Set m_ddClipper = Nothing
    Set m_rm = Nothing
    Set m_dd = Nothing
End Sub

Sub InitScene()
    Set m_rootFrame = m_rm.CreateFrame(Nothing)
    Set m_cameraFrame = m_rm.CreateFrame(m_rootFrame)
    Set m_lightFrame = m_rm.CreateFrame(m_rootFrame)
    Set m_objectFrame = m_rm.CreateFrame(m_rootFrame)
    m_cameraFrame.SetPosition Nothing, 0, 0, -100
    Set m_rmViewport = m_rm.CreateViewport(m_rmDevice, m_cameraFrame, 0, 0, m_width, m_height)
    m_rmViewport.SetBack (1000)
    Set m_light = m_rm.CreateLight(D3DRMLIGHT_DIRECTIONAL, &HFFFFFFFF)
    m_lightFrame.AddLight m_light
    Set m_ambientLight = m_rm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5)   'Shadow
    m_lightFrame.AddLight m_ambientLight
End Sub

Sub AddShapeDX(NumVerts As Integer, VertArray() As D3DVECTOR, SideFaces() As Long, Red, Green, Blue)
    Set m_meshBuilder = m_rm.CreateMeshBuilder()
    Dim NormArray(0) As D3DVECTOR
    m_meshBuilder.AddFaces NumVerts, VertArray, 0, NormArray, SideFaces
    m_meshBuilder.SetColorRGB Red, Green, Blue
    m_objectFrame.AddVisual m_meshBuilder
End Sub

Sub RenderScene()
    On Local Error Resume Next
    m_rmViewport.Clear D3DRMCLEAR_ALL
    m_rmViewport.Render m_rootFrame
    m_rmDevice.Update
    'DoEvents
End Sub

Sub RotateTrackBall(X As Integer, Y As Integer)
    'this function taken from MS engine sample in the VB sample that
    'comes with MS DirectX7 SDK. It works as follows:
        'select point on screen interpret as though selecting
        'point on sphere. as new point is passed when mouse is
        'moved rotate in coresponding direction(s) on sphere.
    Dim delta_x As Single, delta_y As Single
    Dim delta_r As Single, radius As Single, denom As Single, Angle As Single
    ' rotation axis in camcoords, worldcoords, sframecoords
    Dim axisC As D3DVECTOR
    Dim wc As D3DVECTOR
    Dim axisS As D3DVECTOR
    Dim base As D3DVECTOR
    Dim origin As D3DVECTOR
    delta_x = X - m_LastX
    delta_y = Y - m_LastY
    m_LastX = X
    m_LastY = Y
    delta_r = Sqr(delta_x * delta_x + delta_y * delta_y)
    radius = 50
    denom = Sqr(radius * radius + delta_r * delta_r)
    If (delta_r = 0 Or denom = 0) Then Exit Sub
    Angle = (delta_r / denom)
    axisC.X = (-delta_y / delta_r)
    axisC.Y = (-delta_x / delta_r)
    axisC.Z = 0
    m_cameraFrame.Transform wc, axisC
    m_objectFrame.InverseTransform axisS, wc
    m_cameraFrame.Transform wc, origin
    m_objectFrame.InverseTransform base, wc
    axisS.X = axisS.X - base.X
    axisS.Y = axisS.Y - base.Y
    axisS.Z = axisS.Z - base.Z
    m_objectFrame.AddRotation D3DRMCOMBINE_BEFORE, axisS.X, axisS.Y, axisS.Z, Angle
End Sub

Sub RotatePlayer(Angle)
    m_cameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, Angle
End Sub

Sub TiltPlayer(Angle)
    m_cameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, Angle
End Sub

Sub MovePlayerForward(Distance)
    m_cameraFrame.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, Distance
End Sub

Sub MovePlayerUp(Distance)
    m_cameraFrame.AddTranslation D3DRMCOMBINE_BEFORE, 0, Distance / 2, 0
End Sub

Sub PlaceModelinWindow()
    ' This sub adds the model to the DX window in one peice. You cannot remove it again
    ' without clearing the whole window. If you want to make the model move, you
    ' have to run the sub below this one instead...
    InitRM frmMain.DirectX
    InitScene
    
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True Then
            ReDim VertArray(Object(NowOn).VertexCount - 1)
            ReDim SideFaces(Object(NowOn).EdgeCount)
            For Tp = 0 To Object(NowOn).VertexCount - 1
                VertArray(Tp).X = Object(NowOn).Vertex(Tp + 1).X / 8
                VertArray(Tp).Y = -Object(NowOn).Vertex(Tp + 1).Y / 8
                VertArray(Tp).Z = -Object(NowOn).Vertex(Tp + 1).Z / 8
            Next Tp
            FaceON = 1
            For N = 1 To Object(NowOn).FaceCount
                Edges = Object(NowOn).Face(FaceON)
                SideFaces(FaceON - 1) = Edges
                FaceON = FaceON + 1
                For M = 1 To Edges
                    edge = Object(NowOn).Face(FaceON) - 1
                    SideFaces(FaceON - 1) = edge
                    FaceON = FaceON + 1
                Next M
            Next N
            C1 = Object(NowOn).Colour
            Red = C1 And 255
            Green = (C1 And (256 ^ 2 - 256)) / 256
            Blue = (C1 And (256 ^ 3 - 65536)) / (256 ^ 2)
            Red = Red / 255
            Green = Green / 255
            Blue = Blue / 255
            If Red = 0 And Blue = 0 And Green = 0 Then
             Red = 255: Green = 255: Blue = 255
            End If
            AddShapeDX Object(NowOn).VertexCount, VertArray, SideFaces, Red, Green, Blue
            Full = True
        End If
    Next NowOn
End Sub

Sub AnimateDirectX()



Holdin
Dim NowOn As Integer
For NowOn = 1 To cstTotalObjects
    If Object(NowOn).Used = True Then
        JointOn = FindTarget(Object(NowOn).Vertex(1).TargetName)
        If Object(NowOn).Vertex(1).TargetName = "" Then JointOn = 0
        If JointOn <> 99990 Then
            ReDim Rotated(Object(NowOn).VertexCount) As Coord
            ReDim Corner(Object(NowOn).VertexCount, 2)
            For N = 1 To Object(NowOn).VertexCount
                Rotated(N).X = Object(NowOn).Vertex(N).X
                Rotated(N).Y = Object(NowOn).Vertex(N).Y
                Rotated(N).Z = Object(NowOn).Vertex(N).Z
            Next N
            Cx = BaseFrame(JointOn).NewPosit.X
            Cy = BaseFrame(JointOn).NewPosit.Y
            Cz = BaseFrame(JointOn).NewPosit.Z
            Mx = BaseFrame(JointOn).Position.X - BaseFrame(JointOn).NewPosit.X
            My = BaseFrame(JointOn).Position.Y - BaseFrame(JointOn).NewPosit.Y
            Mz = BaseFrame(JointOn).Position.Z - BaseFrame(JointOn).NewPosit.Z
            For N = 1 To Object(NowOn).VertexCount
                Rotated(N).X = Object(NowOn).Vertex(N).X - Mx
                Rotated(N).Y = Object(NowOn).Vertex(N).Y - My
                Rotated(N).Z = Object(NowOn).Vertex(N).Z - Mz
            Next N
            DataOn = (KeepRight(JointOn, 1) * 6) - 5
            If JointOn <> 0 Then

                s1 = 1: S2 = 1:  S3 = 1
                Rotate NowOn, A1, A2, A3, Cx, Cy, Cz, P1, P2, P3, s1, S2, S3, 1
            End If
            
            '#############################################
            
            
            If Object(NowOn).MeshUsed = True Then
                m_objectFrame.DeleteVisual Mesh(NowOn)
                Object(NowOn).MeshUsed = False
            End If
            
            ReDim VertArray(Object(NowOn).VertexCount - 1)
            ReDim SideFaces(Object(NowOn).EdgeCount)
            For Tp = 0 To Object(NowOn).VertexCount - 1
                VertArray(Tp).X = Rotated(Tp + 1).X / 8
                VertArray(Tp).Y = -Rotated(Tp + 1).Y / 8
                VertArray(Tp).Z = -Rotated(Tp + 1).Z / 8
            Next Tp
            FaceON = 1
            For N = 1 To Object(NowOn).FaceCount
                Edges = Object(NowOn).Face(FaceON)
                SideFaces(FaceON - 1) = Edges
                FaceON = FaceON + 1
                For M = 1 To Edges
                    edge = Object(NowOn).Face(FaceON) - 1
                    SideFaces(FaceON - 1) = edge
                    FaceON = FaceON + 1
                Next M
            Next N
            C1 = Object(NowOn).Colour
            Red = C1 And 255
            Green = (C1 And (256 ^ 2 - 256)) / 256
            Blue = (C1 And (256 ^ 3 - 65536)) / (256 ^ 2)
            Red = Red / 255
            Green = Green / 255
            Blue = Blue / 255
            If Red = 0 And Blue = 0 And Green = 0 Then
                Red = 255: Green = 255: Blue = 255
            End If
            
            
            AddAnimatedShapeDX NowOn, Object(NowOn).VertexCount, VertArray, SideFaces, Red, Green, Blue
            
            
            
            '#############################################
        End If
    End If
Next NowOn
RenderScene
End Sub

Sub AddAnimatedShapeDX(NowOn As Integer, NumVerts As Integer, VertArray() As D3DVECTOR, SideFaces() As Long, Red, Green, Blue)
            Set Mesh(NowOn) = m_rm.CreateMeshBuilder()
            Dim NormArray(0) As D3DVECTOR
            Mesh(NowOn).AddFaces NumVerts, VertArray, 0, NormArray, SideFaces
            Mesh(NowOn).SetColorRGB Red, Green, Blue
            m_objectFrame.AddVisual Mesh(NowOn)
            Object(NowOn).MeshUsed = True
End Sub


