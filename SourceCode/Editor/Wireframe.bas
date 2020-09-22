Attribute VB_Name = "ThreeDEngine"
Dim Rotated() As Coord
Dim Ner() As Long
Public SINE(0 To 361) As Double
Public COSINE(0 To 361) As Double
Public Const PI = 3.14159265358979
Public CustomX As Integer, CustomY As Integer

Public Sub DrawGuides(Window As PictureBox, Angle1, Angle2, Angle3, Xf, Yf)
    XX = Xf: YY = Yf
    RotatePoint Angle1, Angle2, Angle3, 40, 0, 0, 0, 0, 0
    X = Rotates.X: Y = Rotates.Y: Z = Rotates.Z
    If frmMain.chk3DOp(2) = True Then
        xx1 = Xf + Int(X * (Zeye / (Zeye - Z)) * Zoom): yy1 = Yf + Int(Y * (Zeye / (Zeye - Z)) * Zoom)
    Else
        xx1 = Xf + Int(X): yy1 = Yf + Int(Y)
    End If
    Window.Line (XX, YY)-(xx1, yy1), RGB(255, 0, 0): Window.Print "X"
    RotatePoint Angle1, Angle2, Angle3, 0, 40, 0, 0, 0, 0
    X = Rotates.X: Y = Rotates.Y: Z = Rotates.Z
    If frmMain.chk3DOp(2) = True Then
        xx1 = Xf + Int(X * (Zeye / (Zeye - Z)) * Zoom): yy1 = Yf + Int(Y * (Zeye / (Zeye - Z)) * Zoom)
    Else
        xx1 = Xf + Int(X): yy1 = Yf + Int(Y)
    End If
    Window.Line (XX, YY)-(xx1, yy1), RGB(0, 255, 0): Window.Print "Y"
    RotatePoint Angle1, Angle2, Angle3, 0, 0, 40, 0, 0, 0
    X = Rotates.X: Y = Rotates.Y: Z = Rotates.Z
    If frmMain.chk3DOp(2) = True Then
        xx1 = Xf + Int(X * (Zeye / (Zeye - Z)) * Zoom): yy1 = Yf + Int(Y * (Zeye / (Zeye - Z)) * Zoom)
    Else
        xx1 = Xf + Int(X): yy1 = Yf + Int(Y)
    End If
    Window.Line (XX, YY)-(xx1, yy1), RGB(0, 0, 255): Window.Print "Z"
End Sub

Public Sub FindNewSkeliton()
    If Animate = True Or sceneEdit.Visible = True Then
        Holdin
    Else
        For N = 1 To cstTotalJoints
            BaseFrame(N).NewPosit.X = BaseFrame(N).Position.X
            BaseFrame(N).NewPosit.Y = BaseFrame(N).Position.Y
            BaseFrame(N).NewPosit.Z = BaseFrame(N).Position.Z
        Next N
    End If

End Sub

Public Sub Draw3DSkeliton(Window As PictureBox, Angle1, Angle2, Angle3, Xf, Yf)
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True Then
        
            X = BaseFrame(N).NewPosit.X
            Y = BaseFrame(N).NewPosit.Y
            Z = BaseFrame(N).NewPosit.Z
            
            RotatePoint Angle1, Angle2, Angle3, X, Y, Z, 0, 0, 0
            X = Rotates.X: Y = Rotates.Y: Z = Rotates.Z
            If frmMain.chk3DOp(2).Value = 1 Then
                XX = Xf + Int(X * (Zeye / (Zeye - Z)) * Zoom)
                YY = Yf + Int(Y * (Zeye / (Zeye - Z)) * Zoom)
            Else
                XX = Xf + X
                YY = Yf + Y
            End If
            HaHa(N, 1) = XX
            HaHa(N, 2) = YY
            
            Tagt = FindTarget(BaseFrame(N).Target)
          '  MsgBox tagt
            If Tagt <> 0 Then
            
                X = BaseFrame(Tagt).NewPosit.X
                Y = BaseFrame(Tagt).NewPosit.Y
                Z = BaseFrame(Tagt).NewPosit.Z
            
                RotatePoint Angle1, Angle2, Angle3, X, Y, Z, 0, 0, 0
                X = Rotates.X: Y = Rotates.Y: Z = Rotates.Z
            
                If frmMain.chk3DOp(2).Value = 1 Then
                    xx2 = Xf + Int(X * (Zeye / (Zeye - Z)) * Zoom)
                    yy2 = Yf + Int(Y * (Zeye / (Zeye - Z)) * Zoom)
                Else
                    xx2 = Xf + X
                    yy2 = Yf + Y
                End If
            
                Window.Line (XX, YY)-(xx2, yy2)
            End If
            Window.Circle (XX, YY), 4
            
            If frmMain.chk3DOp(5) = 1 Then
                temp = Window.ForeColor: Window.ForeColor = vbRed
                Window.Print BaseFrame(N).Name: Window.ForeColor = temp
            End If
            If BaseFrame(N).Selected = True Then Window.Circle (xx1, yy1), 5
        End If
    Next N
End Sub

Public Sub Draw3DBrush(Window As PictureBox, NowOn, Angle1, Angle2, Angle3, ShadeThis, Xf, Yf, Cx, Cy, Cz)
    xof = Window.ScaleWidth / 2
    yof = Window.ScaleHeight / 2
    ReDim Rotated(Object(NowOn).VertexCount)
    ReDim Corner(Object(NowOn).VertexCount, 2)
    
    For N = 1 To Object(NowOn).VertexCount
        Rotated(N).X = Object(NowOn).Vertex(N).X
        Rotated(N).Y = Object(NowOn).Vertex(N).Y
        Rotated(N).Z = Object(NowOn).Vertex(N).Z
    Next N
    
    
    
    If sceneEdit.Visible = True Or Animate = True Then
        For mm = 1 To Object(NowOn).VertexCount
            JointOn = Object(NowOn).Vertex(mm).Target
            Cx = BaseFrame(JointOn).NewPosit.X
            Cy = BaseFrame(JointOn).NewPosit.Y
            Cz = BaseFrame(JointOn).NewPosit.Z
            Mx = BaseFrame(JointOn).Position.X - BaseFrame(JointOn).NewPosit.X
            My = BaseFrame(JointOn).Position.Y - BaseFrame(JointOn).NewPosit.Y
            Mz = BaseFrame(JointOn).Position.Z - BaseFrame(JointOn).NewPosit.Z
            Rotated(mm).X = Object(NowOn).Vertex(mm).X - Mx
            Rotated(mm).Y = Object(NowOn).Vertex(mm).Y - My
            Rotated(mm).Z = Object(NowOn).Vertex(mm).Z - Mz
            DataOn = (KeepRight(JointOn, 1) * 6) - 5
            If JointOn <> 0 And DataOn <> -5 Then
                A1 = BaseFrame(JointOn).Angle.X
                A2 = BaseFrame(JointOn).Angle.Y
                A3 = BaseFrame(JointOn).Angle.Z
                s1 = BaseFrame(JointOn).Scale.X
                S2 = BaseFrame(JointOn).Scale.Y
                S3 = BaseFrame(JointOn).Scale.Z
                Rotate3 NowOn, mm, A1, A2, A3, Cx, Cy, Cz, P1, P2, P3, s1, S2, S3
            End If
            ssss = ssss & mm & " " & JointOn & vbNewLine
        Next mm
        Cx = 0: Cy = 0: Cz = 0
    End If
    
    
    Rotate NowOn, Angle1, Angle2, Angle3, Cx, Cy, Cz, 0, 0, 0, 1, 1, 1, Object(NowOn).VertexCount
    
    For N = 1 To Object(NowOn).VertexCount
            X = Rotated(N).X
            Y = Rotated(N).Y
            Z = Rotated(N).Z
            If frmMain.chk3DOp(2) = 1 Then
                Corner(N, 1) = xof + Int(X * (Zeye / (Zeye - Z)) * Zoom)
                Corner(N, 2) = yof + Int(Y * (Zeye / (Zeye - Z)) * Zoom)
            Else
                Corner(N, 1) = xof + X * Zoom
                Corner(N, 2) = yof + Y * Zoom
            End If
    Next N
    edge = 1
    For N = 1 To Object(NowOn).FaceCount
        EdgeCount = Object(NowOn).Face(edge): edge = edge + 1
        StartOfFace = edge
        Hi1 = Object(NowOn).Face(StartOfFace)
        hi2 = Object(NowOn).Face(StartOfFace + 1)
        hi3 = Object(NowOn).Face(StartOfFace + 2)
        Normal = ((Corner(Hi1, 2) - Corner(hi3, 2)) * (Corner(hi2, 1) - Corner(Hi1, 1)) - (Corner(Hi1, 1) - Corner(hi3, 1)) * (Corner(hi2, 2) - Corner(Hi1, 2)))
        DrawThisFace = False
        If Normal < 0 Then DrawThisFace = True
        If frmMain.chk3DOp(4) = 0 Then DrawThisFace = True
        For M = 1 To EdgeCount
            If M = EdgeCount Then
                Corner1 = Object(NowOn).Face(edge)
                Corner2 = Object(NowOn).Face(StartOfFace): edge = edge + 1
                x1 = Corner(Corner1, 1)
                x2 = Corner(Corner2, 1)
                y1 = Corner(Corner1, 2)
                y2 = Corner(Corner2, 2)
            Else
                Corner1 = Object(NowOn).Face(edge): edge = edge + 1
                Corner2 = Object(NowOn).Face(edge)
                x1 = Corner(Corner1, 1)
                x2 = Corner(Corner2, 1)
                y1 = Corner(Corner1, 2)
                y2 = Corner(Corner2, 2)
            End If
            If DrawThisFace = True Then
                Window.Line (x1, y1)-(x2, y2), Object(NowOn).Colour
            End If
        Next M
    Next N
End Sub


Public Sub SpinSelected(SpinAngle)
    Model.Saved = False
    Dim NowOn As Integer
    xof = frmMain.view.ScaleWidth / 2
    yof = frmMain.view.ScaleHeight / 2
    If frmMain.optGetCenter(0) = True Then
        Cx = FindCenter("X", 0)
        Cy = FindCenter("y", 0)
        Cz = FindCenter("z", 0)
    End If
    If frmMain.optGetCenter(1) = True Then
        Cx = 0
        Cy = 0
        Cz = 0
    End If
    If frmMain.optGetCenter(2) = True Then
        If ViewMode = 1 Then
            Cx = CustomX: Cz = CustomY
        End If
        If ViewMode = 2 Then
            Cx = CustomX: Cy = CustomY
        End If
        If ViewMode = 3 Then
            Cz = CustomX: Cy = CustomY
        End If
    End If
    Cnt = 0
    If frmMain.optGetCenter(5) = True Then
        For N = 1 To cstTotalObjects
            If Object(N).Selected = True Then
                For M = 1 To Object(N).VertexCount
                    If Object(N).Vertex(M).Selected = True Then
                        Cnt = Cnt + 1
                        Cx = Cx + Object(N).Vertex(M).X
                        Cy = Cy + Object(N).Vertex(M).Y
                        Cz = Cz + Object(N).Vertex(M).Z
                    End If
                Next M
            End If
        Next N
    End If
    If Cnt <> 0 Then
        Cx = Cx / Cnt
        Cy = Cy / Cnt
        Cz = Cz / Cnt
    End If
    
    
    frmMain.SBar.Panels(3) = Cx & " " & Cz
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True And Object(NowOn).Selected = True Then
            If frmMain.optGetCenter(4) = True Then
                Cx = FindCenter("X", NowOn)
                Cy = FindCenter("y", NowOn)
                Cz = FindCenter("z", NowOn)
            End If
            If frmMain.optGetCenter(3) = True Then
                JointOn = FindTarget(Object(NowOn).Vertex(1).TargetName)
                Cx = BaseFrame(JointOn).Position.X
                Cy = BaseFrame(JointOn).Position.Y
                Cz = BaseFrame(JointOn).Position.Z
            End If
            ReDim Rotated(Object(NowOn).VertexCount)
            For N = 1 To Object(NowOn).VertexCount
                Rotated(N).X = Object(NowOn).Vertex(N).X
                Rotated(N).Y = Object(NowOn).Vertex(N).Y
                Rotated(N).Z = Object(NowOn).Vertex(N).Z
            Next N
            If ViewMode = 1 Then Rotate NowOn, 0, SpinAngle, 0, Cx, Cy, Cz, 0, 0, 0, 1, 1, 1, Object(NowOn).VertexCount
            If ViewMode = 2 Then Rotate NowOn, 0, 0, SpinAngle, Cx, Cy, Cz, 0, 0, 0, 1, 1, 1, Object(NowOn).VertexCount
            If ViewMode = 3 Then Rotate NowOn, 360 - SpinAngle, 0, 0, Cx, Cy, Cz, 0, 0, 0, 1, 1, 1, Object(NowOn).VertexCount
            For N = 1 To Object(NowOn).VertexCount
                If frmMain.chkVert = 0 Or Object(NowOn).Vertex(N).Selected = True Then
                    Object(NowOn).Vertex(N).X = Rotated(N).X
                    Object(NowOn).Vertex(N).Y = Rotated(N).Y
                    Object(NowOn).Vertex(N).Z = Rotated(N).Z
                End If
            Next N
            FindOutLine NowOn
        End If
    Next NowOn
    For NowOn = 1 To cstTotalJoints
        If BaseFrame(NowOn).Selected = True Then
            ReDim Rotated(1)
            Rotated(1).X = BaseFrame(NowOn).Position.X
            Rotated(1).Y = BaseFrame(NowOn).Position.Y
            Rotated(1).Z = BaseFrame(NowOn).Position.Z
            If ViewMode = 1 Then Rotate NowOn, 0, SpinAngle, 0, Cx, Cy, Cz, 0, 0, 0, 1, 1, 1, 1
            If ViewMode = 2 Then Rotate NowOn, 0, 0, SpinAngle, Cx, Cy, Cz, 0, 0, 0, 1, 1, 1, 1
            If ViewMode = 3 Then Rotate NowOn, 360 - SpinAngle, 0, 0, Cx, Cy, Cz, 0, 0, 0, 1, 1, 1, 1
            BaseFrame(NowOn).Position.X = Rotated(1).X
            BaseFrame(NowOn).Position.Y = Rotated(1).Y
            BaseFrame(NowOn).Position.Z = Rotated(1).Z
        End If
    Next NowOn
    frmMain.DrawModel
End Sub


Public Function Rotate3(NowOn, N, Angle1, Angle2, Angle3, Cx, Cy, Cz, P1, P2, P3, s1, S2, S3)
    tAngle1 = (Angle1 Mod 360)
    tAngle2 = (Angle2 Mod 360)
    tAngle3 = (Angle3 Mod 360)
    If tAngle1 < 0 Then tAngle1 = tAngle1 + 360
    If tAngle2 < 0 Then tAngle2 = tAngle2 + 360
    If tAngle3 < 0 Then tAngle3 = tAngle3 + 360
        X = Rotated(N).X - Cx
        Y = Rotated(N).Y - Cy
        Z = Rotated(N).Z - Cz
        XRotated = X
        YRotated = COSINE(tAngle1) * Y - SINE(tAngle1) * Z
        ZRotated = SINE(tAngle1) * Y + COSINE(tAngle1) * Z
        X = XRotated
        Y = YRotated
        Z = ZRotated
        XRotated = COSINE(tAngle2) * X - SINE(tAngle2) * Z
        YRotated = Y
        ZRotated = SINE(tAngle2) * X + COSINE(tAngle2) * Z
        X = XRotated
        Y = YRotated
        Z = ZRotated
        XRotated = COSINE(tAngle3) * X - SINE(tAngle3) * Y
        YRotated = SINE(tAngle3) * X + COSINE(tAngle3) * Y
        ZRotated = Z
        Rotated(N).X = (XRotated * s1) + Cx + P1
        Rotated(N).Y = (YRotated * S2) + Cy + P2
        Rotated(N).Z = (ZRotated * S3) + Cz + P3
End Function


Public Function Rotate(NowOn, Angle1, Angle2, Angle3, Cx, Cy, Cz, P1, P2, P3, s1, S2, S3, Points2Spin)
    tAngle1 = (Angle1 Mod 360)
    tAngle2 = (Angle2 Mod 360)
    tAngle3 = (Angle3 Mod 360)
    If tAngle1 < 0 Then tAngle1 = tAngle1 + 360
    If tAngle2 < 0 Then tAngle2 = tAngle2 + 360
    If tAngle3 < 0 Then tAngle3 = tAngle3 + 360
    For N = 1 To Points2Spin
        X = Rotated(N).X - Cx
        Y = Rotated(N).Y - Cy
        Z = Rotated(N).Z - Cz
        XRotated = X
        YRotated = COSINE(tAngle1) * Y - SINE(tAngle1) * Z
        ZRotated = SINE(tAngle1) * Y + COSINE(tAngle1) * Z
        X = XRotated
        Y = YRotated
        Z = ZRotated
        XRotated = COSINE(tAngle2) * X - SINE(tAngle2) * Z
        YRotated = Y
        ZRotated = SINE(tAngle2) * X + COSINE(tAngle2) * Z
        X = XRotated
        Y = YRotated
        Z = ZRotated
        XRotated = COSINE(tAngle3) * X - SINE(tAngle3) * Y
        YRotated = SINE(tAngle3) * X + COSINE(tAngle3) * Y
        ZRotated = Z
        Rotated(N).X = (XRotated * s1) + Cx + P1
        Rotated(N).Y = (YRotated * S2) + Cy + P2
        Rotated(N).Z = (ZRotated * S3) + Cz + P3
    Next N
End Function


Sub Make_LookUp()
    For i = 0 To 361
        SINE(i) = Sin(i / 180 * PI)
        COSINE(i) = Cos(i / 180 * PI)
    Next
End Sub


Sub DrawPolygon(Window, FaceUsed, cola, EdgeCount)
    If FaceUsed = 0 Then cola = RGB(220, 220, 220)
    Dim nn As Integer
    ReDim StartStop(Window.ScaleHeight, 1 To 3) As Long
    For nn = 1 To EdgeCount
        If nn = EdgeCount Then
            x1 = Ner(nn, 1)
            y1 = Ner(nn, 2)
            x2 = Ner(1, 1)
            y2 = Ner(1, 2)
            GoSub DrawLine
        Else
            x1 = Ner(nn, 1)
            y1 = Ner(nn, 2)
            x2 = Ner(nn + 1, 1)
            y2 = Ner(nn + 1, 2)
            GoSub DrawLine
        End If
    Next nn
    Exit Sub
DrawLine:
    mox = x2 - x1: moy = y2 - y1
    If moy <> 0 Then
        Disd = mox / moy
    Else
        Return
    End If
    If y1 < y2 Then Negpos = 1:  Disd = Disd
    If y1 > y2 Then Negpos = -1: Disd = -Disd
    Tag = y1: XX = x1
    For N = y1 To y2 Step Negpos
        If N > -1 And N < Window.ScaleHeight Then
            If StartStop(N, 3) = 1 Then
                StartStop(N, 3) = 1
                If N / 2 = Int(N / 2) Then Window.Line (StartStop(N, 1), N)-(XX, N), cola
            Else
                StartStop(N, 3) = 1
                StartStop(N, 1) = XX
            End If
        End If
        XX = XX + Disd
    Next N
    Return
End Sub

Sub RotatePoint(Angle1, Angle2, Angle3, X, Y, Z, Mx, My, Mz)
    tAngle1 = (Angle1 Mod 360)
    tAngle2 = (Angle2 Mod 360)
    tAngle3 = (Angle3 Mod 360)
    If tAngle1 < 0 Then tAngle1 = tAngle1 + 360
    If tAngle2 < 0 Then tAngle2 = tAngle2 + 360
    If tAngle3 < 0 Then tAngle3 = tAngle3 + 360
    X = X - Mx
    Y = Y - My
    Z = Z - Mz
    XRotated = X
    YRotated = COSINE(tAngle1) * Y - SINE(tAngle1) * Z
    ZRotated = SINE(tAngle1) * Y + COSINE(tAngle1) * Z
    X = XRotated
    Y = YRotated
    Z = ZRotated
    XRotated = COSINE(tAngle2) * X - SINE(tAngle2) * Z
    YRotated = Y
    ZRotated = SINE(tAngle2) * X + COSINE(tAngle2) * Z
    X = XRotated
    Y = YRotated
    Z = ZRotated
    XRotated = COSINE(tAngle3) * X - SINE(tAngle3) * Y
    YRotated = SINE(tAngle3) * X + COSINE(tAngle3) * Y
    ZRotated = Z
    Rotates.X = XRotated + Mx
    Rotates.Y = YRotated + My
    Rotates.Z = ZRotated + Mz
End Sub

Public Sub Holdin()
    For N = 1 To cstTotalJoints
            BaseFrame(N).Angle.X = 0
            BaseFrame(N).Angle.Y = 0
            BaseFrame(N).Angle.Z = 0
    Next N
    If sceneEdit.Data.Count = 1 Then Exit Sub
    For nn = 1 To CountJoints
        N = RememberW(nn)
        If BaseFrame(N).Used = True And KeepRight(N, 2) <> 0 Then
            Target = FindTarget(BaseFrame(N).Target)
                fool = (KeepRight(N, 2) * 9) - 8
                aa1 = sceneEdit.Data(fool)
                aa2 = sceneEdit.Data(fool + 1)
                aa3 = sceneEdit.Data(fool + 2)
                pp1 = sceneEdit.Data(fool + 3)
                pp2 = sceneEdit.Data(fool + 4)
                pp3 = sceneEdit.Data(fool + 5)
                ss1 = 1 + (sceneEdit.Data(fool + 6) * 0.01)
                ss2 = 1 + (sceneEdit.Data(fool + 7) * 0.01)
                ss3 = 1 + (sceneEdit.Data(fool + 8) * 0.01)
                Cx = BaseFrame(Target).NewPosit.X
                Cy = BaseFrame(Target).NewPosit.Y
                Cz = BaseFrame(Target).NewPosit.Z
                sAngle1 = BaseFrame(Target).Angle.X
                sAngle2 = BaseFrame(Target).Angle.Y
                sAngle3 = BaseFrame(Target).Angle.Z
                X = BaseFrame(Target).NewPosit.X + BaseFrame(N).Position.X - BaseFrame(Target).Position.X + pp1
                Y = BaseFrame(Target).NewPosit.Y + BaseFrame(N).Position.Y - BaseFrame(Target).Position.Y + pp2
                Z = BaseFrame(Target).NewPosit.Z + BaseFrame(N).Position.Z - BaseFrame(Target).Position.Z + pp3
                BaseFrame(N).Angle.X = aa1 + sAngle1
                BaseFrame(N).Angle.Y = aa2 + sAngle2
                BaseFrame(N).Angle.Z = aa3 + sAngle3
                sRotate sAngle1, sAngle2, sAngle3, X, Y, Z, Cx, Cy, Cz
                BaseFrame(N).NewPosit.X = Rotates.X
                BaseFrame(N).NewPosit.Y = Rotates.Y
                BaseFrame(N).NewPosit.Z = Rotates.Z
                BaseFrame(N).Scale.X = ss1
                BaseFrame(N).Scale.Y = ss2
                BaseFrame(N).Scale.Z = ss3
        End If
    Next nn
End Sub

Public Sub sRotate(tAngle1, tAngle2, tAngle3, X, Y, Z, Cx, Cy, Cz)
    tAngle1 = (tAngle1 Mod 360)
    tAngle2 = (tAngle2 Mod 360)
    tAngle3 = (tAngle3 Mod 360)
    If tAngle1 < 0 Then tAngle1 = tAngle1 + 360
    If tAngle2 < 0 Then tAngle2 = tAngle2 + 360
    If tAngle3 < 0 Then tAngle3 = tAngle3 + 360
    X = X - Cx
    Y = Y - Cy
    Z = Z - Cz
    XRotated = X
    YRotated = COSINE(tAngle1) * Y - SINE(tAngle1) * Z
    ZRotated = SINE(tAngle1) * Y + COSINE(tAngle1) * Z
    X = XRotated
    Y = YRotated
    Z = ZRotated
    XRotated = COSINE(tAngle2) * X - SINE(tAngle2) * Z
    YRotated = Y
    ZRotated = SINE(tAngle2) * X + COSINE(tAngle2) * Z
    X = XRotated
    Y = YRotated
    Z = ZRotated
    XRotated = COSINE(tAngle3) * X - SINE(tAngle3) * Y
    YRotated = SINE(tAngle3) * X + COSINE(tAngle3) * Y
    ZRotated = Z
    Rotates.X = XRotated + Cx
    Rotates.Y = YRotated + Cy
    Rotates.Z = ZRotated + Cz
End Sub

