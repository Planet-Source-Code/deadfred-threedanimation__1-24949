Attribute VB_Name = "SelectDeselect"
Option Explicit

Public Function ClickSelect(X, Y, Shift) As Boolean
    Dim xof As Single, yof As Single, NowOn As Integer, N As Integer, XX As Integer, YY As Integer, M As Integer
    ClickSelect = False
    If frmMain.chkSelDesel = 1 Then Exit Function
    xof = (frmMain.View.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.View.ScaleHeight / 2) - frmMain.UDBar
    If frmMain.chkObjects = 1 Then
        For NowOn = 1 To cstTotalObjects
            If Object(NowOn).Used = True Then
            
            
            
            
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            For N = 1 To Object(NowOn).VertexCount
                                If ViewMode = 1 Then
                                    XX = (Object(NowOn).Vertex(N).X * Zoom) + xof
                                    YY = (Object(NowOn).Vertex(N).Z * Zoom) + yof
                                End If
                                If ViewMode = 2 Then
                                    XX = (Object(NowOn).Vertex(N).X * Zoom) + xof
                                    YY = (Object(NowOn).Vertex(N).Y * Zoom) + yof
                                End If
                                If ViewMode = 3 Then
                                    XX = (Object(NowOn).Vertex(N).Z * Zoom) + xof
                                    YY = (Object(NowOn).Vertex(N).Y * Zoom) + yof
                                End If
                                If Almostt((X), (Y), XX, YY, 5 * Zoom) = True Then
                                    If Shift = 0 And frmMain.chkDeselect = 0 Then Object(NowOn).Selected = True
                                    If Shift = 0 And frmMain.chkDeselect = 1 Then DeselectAll: Object(NowOn).Selected = True
                                    If Shift = 1 Then Object(NowOn).Selected = True
                                    If Shift = 2 Then Object(NowOn).Selected = False
                                    For M = 1 To cstTotalObjects
                                        If Object(M).GroupCount <> 0 And Object(NowOn).GroupCount <> 0 Then
                                            If Object(M).Group(Object(M).GroupCount) = Object(NowOn).Group(Object(NowOn).GroupCount) Then
                                                Object(M).Selected = True
                                            End If
                                        End If
                                    Next M
                                    ClickSelect = True
                                End If
                            Next N
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            End If
        Next NowOn
    End If
    If frmMain.chkJoints = 1 Then
        For NowOn = 1 To cstTotalJoints
            If BaseFrame(NowOn).Used = True Then
                If ViewMode = 1 Then
                    XX = (BaseFrame(NowOn).Position.X * Zoom) + xof
                    YY = (BaseFrame(NowOn).Position.Z * Zoom) + yof
                End If
                If ViewMode = 2 Then
                    XX = (BaseFrame(NowOn).Position.X * Zoom) + xof
                    YY = (BaseFrame(NowOn).Position.Y * Zoom) + yof
                End If
                If ViewMode = 3 Then
                    XX = (BaseFrame(NowOn).Position.Z * Zoom) + xof
                    YY = (BaseFrame(NowOn).Position.Y * Zoom) + yof
                End If
                If Almostt((X), (Y), XX, YY, 5 * Zoom) = True Then
                    If Shift = 0 And frmMain.chkDeselect = 0 Then BaseFrame(NowOn).Selected = True
                    If Shift = 0 And frmMain.chkDeselect = 1 Then
                        BaseFrame(NowOn).Selected = True
                        ClickSelect = True
                        If frmMain.chkAttchOb = 1 Then
                            For N = 1 To cstTotalObjects
                                If Object(N).Used = True Then
                                    For M = 1 To Object(N).VertexCount
                                        If Object(N).Vertex(M).TargetName = BaseFrame(NowOn).Name Then
                                            Object(N).Selected = True
                                            Exit For
                                        End If
                                    Next M
                                End If
                            Next N
                        End If
                    End If
                    If Shift = 1 Then BaseFrame(NowOn).Selected = True
                    If Shift = 2 Then BaseFrame(NowOn).Selected = False
                End If
            End If
        Next NowOn
    End If
    frmMain.View.DrawStyle = 0
    frmMain.DrawModel
End Function

Public Sub BoxBandSelect(x1, y1, x2, y2, Shift)
    Dim temp As Integer, xof As Integer, yof As Integer, NowOn As Integer, Count As Integer
    Dim N As Integer, X As Integer, Y As Integer, M As Integer
    If frmMain.chkSelDesel = 1 Then Exit Sub
    If x1 > x2 Then temp = x1: x1 = x2: x2 = temp
    If y1 > y2 Then temp = y1: y1 = y2: y2 = temp
    If Shift = 0 Then DeselectAll
    xof = (frmMain.View.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.View.ScaleHeight / 2) - frmMain.UDBar
    If frmMain.chkObjects = 1 Then
        For NowOn = 1 To cstTotalObjects
            If Object(NowOn).Used = True Then
                Count = 0
                For N = 1 To Object(NowOn).VertexCount
                    If ViewMode = 1 Then
                        X = (Object(NowOn).Vertex(N).X * Zoom) + xof
                        Y = (Object(NowOn).Vertex(N).Z * Zoom) + yof
                    End If
                    If ViewMode = 2 Then
                        X = (Object(NowOn).Vertex(N).X * Zoom) + xof
                        Y = (Object(NowOn).Vertex(N).Y * Zoom) + yof
                    End If
                    If ViewMode = 3 Then
                        X = (Object(NowOn).Vertex(N).Z * Zoom) + xof
                        Y = (Object(NowOn).Vertex(N).Y * Zoom) + yof
                    End If
                    If (X > x1 And X < x2) And (Y > y1 And Y < y2) Then
                        Count = Count + 1
                        If frmMain.chkTotalSelect = 0 Then Object(NowOn).Selected = True
                        For M = 1 To cstTotalObjects
                            If Object(M).GroupCount <> 0 And Object(NowOn).GroupCount <> 0 Then
                                If Object(M).Group(Object(M).GroupCount) = Object(NowOn).Group(Object(NowOn).GroupCount) Then
                                    Object(M).Selected = True
                                End If
                            End If
                        Next M
                    End If
                Next N
                If Count = Object(NowOn).VertexCount Then
                    Object(NowOn).Selected = True
                    For M = 1 To cstTotalObjects
                        If Object(M).GroupCount <> 0 And Object(NowOn).GroupCount <> 0 Then
                            If Object(M).Group(Object(M).GroupCount) = Object(NowOn).Group(Object(NowOn).GroupCount) Then
                                Object(M).Selected = True
                            End If
                        End If
                    Next M
                End If
            End If
        Next NowOn
    End If
    If frmMain.chkJoints = 1 Then
        For NowOn = 1 To cstTotalJoints
            If BaseFrame(NowOn).Used = True Then
                If ViewMode = 1 Then
                    X = (BaseFrame(NowOn).Position.X * Zoom) + xof
                    Y = (BaseFrame(NowOn).Position.Z * Zoom) + yof
                End If
                If ViewMode = 2 Then
                    X = (BaseFrame(NowOn).Position.X * Zoom) + xof
                    Y = (BaseFrame(NowOn).Position.Y * Zoom) + yof
                End If
                If ViewMode = 3 Then
                    X = (BaseFrame(NowOn).Position.Z * Zoom) + xof
                    Y = (BaseFrame(NowOn).Position.Y * Zoom) + yof
                End If
                If (X > x1 And X < x2) And (Y > y1 And Y < y2) Then
                    If Shift = 0 And frmMain.chkDeselect = 0 Then BaseFrame(NowOn).Selected = True
                    If Shift = 0 And frmMain.chkDeselect = 1 Then BaseFrame(NowOn).Selected = True
                    If Shift = 1 Then BaseFrame(NowOn).Selected = True
                    If Shift = 2 Then BaseFrame(NowOn).Selected = False
                End If
            End If
        Next NowOn
    End If
    
    FindObjectOutline
    frmMain.DrawModel
End Sub

Public Sub SelectAll()
    Dim NowOn As Integer
    If frmMain.chkSelDesel = 0 And frmMain.chkObjects = 1 Then
        For NowOn = 1 To cstTotalObjects
            Object(NowOn).Selected = True
        Next NowOn
    End If
    If frmMain.chkSelDesel = 0 And frmMain.chkJoints = 1 Then
        For NowOn = 1 To cstTotalJoints
            BaseFrame(NowOn).Selected = True
        Next NowOn
    End If
    
    
    
End Sub

Public Sub DeselectAll()
    Dim NowOn As Integer
    If frmMain.chkSelDesel = 1 Then Exit Sub
    frmMain.OpEdit(1).Enabled = False
    For NowOn = 1 To cstTotalObjects
        Object(NowOn).Selected = False
    Next NowOn
    For NowOn = 1 To cstTotalJoints
        BaseFrame(NowOn).Selected = False
    Next NowOn
    Model.Max.X = 0:   Model.Max.Y = 0:   Model.Max.Z = 0
    Model.Min.X = 0:    Model.Min.Y = 0:    Model.Min.Z = 0
End Sub

Public Sub GroupSelected()
    Dim NowOn As Integer, bFirst As Boolean, GroupCode As Integer, Count As Integer
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True And Object(NowOn).Selected = True Then
            If bFirst = False Then
                GroupCode = NowOn
                bFirst = True
            End If
            Object(NowOn).GroupCount = Object(NowOn).GroupCount + 1
            ReDim Preserve Object(NowOn).Group(Object(NowOn).GroupCount)
            Object(NowOn).Group(Object(NowOn).GroupCount) = GroupCode
            Count = Count + 1
        End If
    Next NowOn
    frmMain.Sbar.Panels(3) = Count & " objects grouped"
End Sub

Public Sub UnGroupSelected()
    Dim NowOn As Integer, bFirst As Boolean, GroupCode As Integer, Count As Integer
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True And Object(NowOn).Selected = True And Object(NowOn).GroupCount <> 0 Then
            Object(NowOn).GroupCount = Object(NowOn).GroupCount - 1
            ReDim Preserve Object(NowOn).Group(Object(NowOn).GroupCount)
            Count = Count + 1
        End If
    Next NowOn
    frmMain.Sbar.Panels(3) = Count & " objects ungrouped"
End Sub


Public Sub FindOutLine(NowOn)
    Dim N As Integer, x1 As Integer, y1 As Integer, z1 As Integer
    For N = 1 To Object(NowOn).VertexCount
            x1 = Object(NowOn).Vertex(N).X
            y1 = Object(NowOn).Vertex(N).Y
            z1 = Object(NowOn).Vertex(N).Z
            If N = 1 Then
                Object(NowOn).Min.X = x1:       Object(NowOn).Max.X = x1
                Object(NowOn).Min.Y = y1:       Object(NowOn).Max.Y = y1
                Object(NowOn).Min.Z = z1:       Object(NowOn).Max.Z = z1
            End If
            If y1 < Object(NowOn).Min.Y Then Object(NowOn).Min.Y = y1
            If y1 > Object(NowOn).Max.Y Then Object(NowOn).Max.Y = y1
            If x1 < Object(NowOn).Min.X Then Object(NowOn).Min.X = x1
            If x1 > Object(NowOn).Max.X Then Object(NowOn).Max.X = x1
            If z1 < Object(NowOn).Min.Z Then Object(NowOn).Min.Z = z1
            If z1 > Object(NowOn).Max.Z Then Object(NowOn).Max.Z = z1
    Next N
End Sub

Public Sub FindObjectOutline()
    Dim FirstTime As Boolean, NowOn As Integer, x1 As Integer, y1 As Integer, z1 As Integer
    Dim x2 As Integer, y2 As Integer, z2 As Integer, xof As Integer, yof As Integer
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Selected = True Then
            If FirstTime = False Then
                x1 = Object(NowOn).Max.X
                y1 = Object(NowOn).Max.Y
                z1 = Object(NowOn).Max.Z
                x2 = Object(NowOn).Min.X
                y2 = Object(NowOn).Min.Y
                z2 = Object(NowOn).Min.Z
                FirstTime = True
            End If
            If Object(NowOn).Max.X > x1 Then x1 = Object(NowOn).Max.X
            If Object(NowOn).Max.Y > y1 Then y1 = Object(NowOn).Max.Y
            If Object(NowOn).Max.Z > z1 Then z1 = Object(NowOn).Max.Z
            If Object(NowOn).Min.X < x2 Then x2 = Object(NowOn).Min.X
            If Object(NowOn).Min.Y < y2 Then y2 = Object(NowOn).Min.Y
            If Object(NowOn).Min.Z < z2 Then z2 = Object(NowOn).Min.Z
        End If
    Next NowOn
    For NowOn = 1 To cstTotalJoints
        If BaseFrame(NowOn).Selected = True Then
            If FirstTime = False Then
                x1 = BaseFrame(NowOn).Position.X + 5
                y1 = BaseFrame(NowOn).Position.Y + 5
                z1 = BaseFrame(NowOn).Position.Z + 5
                x2 = BaseFrame(NowOn).Position.X - 5
                y2 = BaseFrame(NowOn).Position.Y - 5
                z2 = BaseFrame(NowOn).Position.Z - 5
                FirstTime = True
            End If
            If BaseFrame(NowOn).Position.X + 5 > x1 Then x1 = BaseFrame(NowOn).Position.X + 5
            If BaseFrame(NowOn).Position.Y + 5 > y1 Then y1 = BaseFrame(NowOn).Position.Y + 5
            If BaseFrame(NowOn).Position.Z + 5 > z1 Then z1 = BaseFrame(NowOn).Position.Z + 5
            If BaseFrame(NowOn).Position.X - 5 < x2 Then x2 = BaseFrame(NowOn).Position.X - 5
            If BaseFrame(NowOn).Position.Y - 5 < y2 Then y2 = BaseFrame(NowOn).Position.Y - 5
            If BaseFrame(NowOn).Position.Z - 5 < z2 Then z2 = BaseFrame(NowOn).Position.Z - 5
        End If
    Next NowOn
    If ViewMode = 1 Then
        Model.Max.X = x1 + xof:        Model.Min.X = x2 + xof
        Model.Max.Y = z1 + yof:        Model.Min.Y = z2 + yof
    End If
    If ViewMode = 2 Then
        Model.Max.X = x1 + xof:        Model.Min.X = x2 + xof
        Model.Max.Y = y1 + yof:        Model.Min.Y = y2 + yof
    End If
    If ViewMode = 3 Then
        Model.Max.X = z1 + xof:        Model.Min.X = z2 + xof
        Model.Max.Y = y1 + yof:        Model.Min.Y = y2 + yof
    End If
End Sub

Public Sub MoveSelected(X, Y)
    Dim NowOn As Integer, Ox As Integer, Oy As Integer, Oz As Integer
    Dim N As Integer, M As Integer
    X = Snaped(X)
    Y = Snaped(Y)
    Model.Saved = False
    If ViewMode = 1 Then Ox = X: Oy = 0: Oz = Y
    If ViewMode = 2 Then Ox = X: Oy = Y: Oz = 0
    If ViewMode = 3 Then Ox = 0: Oy = Y: Oz = X
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True And Object(NowOn).Selected = True Then
            For N = 1 To Object(NowOn).VertexCount
                If frmMain.chkVert = 0 Or Object(NowOn).Vertex(N).Selected = True Then
                    Object(NowOn).Vertex(N).X = CInt(Object(NowOn).Vertex(N).X + (Ox / Zoom))
                    Object(NowOn).Vertex(N).Y = CInt(Object(NowOn).Vertex(N).Y + (Oy / Zoom))
                    Object(NowOn).Vertex(N).Z = CInt(Object(NowOn).Vertex(N).Z + (Oz / Zoom))
                End If
            Next N
            FindOutLine (NowOn)
        End If
    Next NowOn
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Selected = True Then
            BaseFrame(N).Position.X = CInt(BaseFrame(N).Position.X + (Ox / Zoom))
            BaseFrame(N).Position.Y = CInt(BaseFrame(N).Position.Y + (Oy / Zoom))
            BaseFrame(N).Position.Z = CInt(BaseFrame(N).Position.Z + (Oz / Zoom))
        End If
    Next N
    frmMain.DrawModel
End Sub

Public Sub DeleteSelected()
    Dim NowOn As Integer, Num As Integer, XX As Integer, YY As Integer, X As Integer, Y As Integer
    Dim N As Integer, M As Integer, JointOn As String
    If ViewMode > 3 Then Exit Sub
    

    
    
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True And Object(NowOn).Selected = True Then
            Object(NowOn).Used = False
            Object(NowOn).Selected = False
            FindOutLine (NowOn)
        End If
    Next NowOn
    
    If Val(frmMain.CurrentSideBar) = 3 Then
        Num = Val(Mid(frmMain.Joints.SelectedItem.Key, 6, 3))
        BaseFrame(Num).Used = False
        BaseFrame(Num).Selected = False
        If frmMain.Joints.Nodes(frmMain.Joints.SelectedItem.Key).Tag = 1 Then
            Exit Sub
        End If
        JointOn = frmMain.Joints.Nodes(frmMain.Joints.SelectedItem.Key).FullPath
        For N = 1 To cstTotalJoints
            If BaseFrame(N).Used = True Then
                If Mid(frmMain.Joints.Nodes(BaseFrame(N).Target).FullPath & "\" & BaseFrame(N).Name, 1, Len(JointOn)) = JointOn Then
                    BaseFrame(N).Used = False
                    BaseFrame(N).Selected = False
                End If
            End If
        Next N
        frmMain.Joints.Nodes.Remove (frmMain.Joints.SelectedItem.Key)
    Else
    
        For M = 1 To cstTotalJoints
            If BaseFrame(M).Selected = True And BaseFrame(M).Used = True Then
               
               
                JointOn = frmMain.Joints.Nodes(BaseFrame(M).Key).FullPath
                For N = 1 To cstTotalJoints
                    If BaseFrame(N).Used = True Then
                        If Mid(frmMain.Joints.Nodes(BaseFrame(N).Target).FullPath & "\" & BaseFrame(N).Name, 1, Len(JointOn)) = JointOn Then
                            BaseFrame(N).Used = False
                        End If
                    End If
                Next N
               
                frmMain.Joints.Nodes.Remove (BaseFrame(M).Key)
            
            End If
        Next M
    End If

FindObjectOutline

End Sub

Public Sub ScaleObject(Xscale, Yscale, Zscale, Mode)
    Dim NowOn As Integer, N As Integer, XX As Integer, YY As Integer, ZZ As Integer
    Model.Saved = False
    For NowOn = 1 To cstTotalObjects
        If ViewMode = 1 Then
            XX = (Model.Max.X + Model.Min.X) / 2
            ZZ = (Model.Max.Y + Model.Min.Y) / 2
        ElseIf ViewMode = 2 Then
            XX = (Model.Max.X + Model.Min.X) / 2
            YY = (Model.Max.Y + Model.Min.Y) / 2
        Else
            ZZ = (Model.Max.X + Model.Min.X) / 2
            YY = (Model.Max.Y + Model.Min.Y) / 2
        End If
        
        If Object(NowOn).Selected = True Then
            For N = 1 To Object(NowOn).VertexCount
                If frmMain.chkVert = 0 Or Object(NowOn).Vertex(N).Selected = True Then
                    Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X - XX
                    Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X * Xscale
                    Object(NowOn).Vertex(N).X = Object(NowOn).Vertex(N).X + XX
                    Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y - YY
                    Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y * Yscale
                    Object(NowOn).Vertex(N).Y = Object(NowOn).Vertex(N).Y + YY
                    Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z - ZZ
                    Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z * Zscale
                    Object(NowOn).Vertex(N).Z = Object(NowOn).Vertex(N).Z + ZZ
                End If
            Next N
            FindOutLine (NowOn)
        End If
    Next NowOn
End Sub


Public Sub DeselectAllVertecies()
    Dim NowOn As Integer, N As Integer
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Selected = True Then
            For N = 1 To Object(NowOn).VertexCount
                Object(NowOn).Vertex(N).Selected = False
            Next N
        End If
    Next NowOn
End Sub

Public Sub DeleteVertex(NowOn, Vertex)
    Dim N As Integer, EdgeOn As Integer, Vert As Integer, M As Integer, FaceCount As Integer
    Dim L As Integer
    For N = Vertex To Object(NowOn).VertexCount - 1
        Object(NowOn).Vertex(N) = Object(NowOn).Vertex(N + 1)
    Next N
    Object(NowOn).VertexCount = Object(NowOn).VertexCount - 1
    ReDim Preserve Object(NowOn).Vertex(Object(NowOn).VertexCount) As VertexDis
    EdgeOn = 1
    For M = 1 To Object(NowOn).FaceCount
        FaceCount = Object(NowOn).Face(EdgeOn): EdgeOn = EdgeOn + 1
        For L = 1 To FaceCount
            Vert = Object(NowOn).Face(EdgeOn)
            If Vert > Vertex Then
                Object(NowOn).Face(EdgeOn) = Object(NowOn).Face(EdgeOn) - 1
            End If
            EdgeOn = EdgeOn + 1
        Next L
    Next M
End Sub

Public Sub DeleteFace(NowOn, Face)
    Dim FaceON As Integer, N As Integer, M As Integer, edge As Integer, Edges As Integer
    Dim NewON As Integer, NewSHAPE() As Integer, NewFace As Integer, NewFaceOn As Integer
    FaceON = 1
    NewON = 1
    NewFaceOn = 1
    For N = 1 To Object(NowOn).FaceCount
        Edges = Object(NowOn).Face(FaceON)
        If N <> Face Then
            ReDim Preserve NewSHAPE(FaceON) As Integer
            NewSHAPE(NewFaceOn) = Edges
            NewFaceOn = NewFaceOn + 1
        End If
        FaceON = FaceON + 1
        For M = 1 To Edges
            edge = Object(NowOn).Face(FaceON)
            If N <> Face Then
                ReDim Preserve NewSHAPE(NewFaceOn) As Integer
                NewSHAPE(NewFaceOn) = edge
                NewFaceOn = NewFaceOn + 1
            End If
            FaceON = FaceON + 1
        Next M
    Next N
    Object(NowOn).EdgeCount = NewFaceOn - 1
    Object(NowOn).FaceCount = Object(NowOn).FaceCount - 1
    ReDim Object(NowOn).Face(Object(NowOn).EdgeCount) As Integer
    For N = 1 To Object(NowOn).EdgeCount
        Object(NowOn).Face(N) = NewSHAPE(N)
    Next N
End Sub

Public Sub ConvertThisToTriangles(NowOn)
    Dim FaceON As Integer, N As Integer, M As Integer, edge As Integer, Edges As Integer
    Dim NewON As Integer, NewSHAPE() As Integer, NewFace As Integer
    FaceON = 1
    NewON = 1
    For N = 1 To Object(NowOn).FaceCount
        Edges = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
        ReDim Hold(Edges) As Integer
        For M = 1 To Edges
            edge = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
            Hold(M) = edge
        Next M
        For M = 2 To Edges - 1
            ReDim Preserve NewSHAPE(NewON + 4) As Integer
            NewSHAPE(NewON) = 3: NewON = NewON + 1
            NewSHAPE(NewON) = Hold(1): NewON = NewON + 1
            NewSHAPE(NewON) = Hold(M): NewON = NewON + 1
            NewSHAPE(NewON) = Hold(M + 1): NewON = NewON + 1
            NewFace = NewFace + 1
        Next M
    Next N
    Object(NowOn).EdgeCount = NewON
    Object(NowOn).FaceCount = NewFace
    ReDim Object(NowOn).Face(Object(NowOn).EdgeCount) As Integer
    For N = 1 To Object(NowOn).EdgeCount
        Object(NowOn).Face(N) = NewSHAPE(N)
    Next N
End Sub



Public Sub BendFace(NowOn, Face, MidFace)
    Dim FaceON As Integer, N As Integer, M As Integer, edge As Integer, Edges As Integer, xof As Integer
    Dim NewON As Integer, NewSHAPE() As Integer, NewFace As Integer, NewFaceOn As Integer, yof As Integer
    Dim MidX As Integer, MidY As Integer, MidZ As Integer
    xof = (frmMain.View.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.View.ScaleHeight / 2) - frmMain.UDBar
    FaceON = 1:    NewON = 1:    NewFaceOn = 1
    For N = 1 To Object(NowOn).FaceCount
        Edges = Object(NowOn).Face(FaceON)
        If N <> Face Then
            ReDim Preserve NewSHAPE(FaceON) As Integer
            NewSHAPE(NewFaceOn) = Edges
            NewFaceOn = NewFaceOn + 1
        Else
            ReDim ThaNewFace(Edges + 1) As Integer
            ThaNewFace(1) = Edges
        End If
        FaceON = FaceON + 1
        For M = 1 To Edges
            edge = Object(NowOn).Face(FaceON)
            If N <> Face Then
                ReDim Preserve NewSHAPE(NewFaceOn) As Integer
                NewSHAPE(NewFaceOn) = edge: NewFaceOn = NewFaceOn + 1
            Else
                ThaNewFace(M + 1) = edge
            End If
            FaceON = FaceON + 1
        Next M
    Next N
    
    Object(NowOn).EdgeCount = NewFaceOn - 1
    Object(NowOn).FaceCount = Object(NowOn).FaceCount - 1
    ReDim Object(NowOn).Face(Object(NowOn).EdgeCount) As Integer
    
    For N = 1 To Object(NowOn).EdgeCount
        Object(NowOn).Face(N) = NewSHAPE(N)
        If N <> 1 Then
            MidX = MidX + Object(NowOn).Vertex(Object(NowOn).Face(N)).X
            MidY = MidY + Object(NowOn).Vertex(Object(NowOn).Face(N)).Y
            MidZ = MidZ + Object(NowOn).Vertex(Object(NowOn).Face(N)).Z
        End If
    Next N
    MidX = Snaped(MidX / (N - 1))
    MidY = Snaped(MidY / (N - 1))
    MidZ = Snaped(MidZ / (N - 1))
    
    Object(NowOn).VertexCount = Object(NowOn).VertexCount + 1
    ReDim Preserve Object(NowOn).Vertex(Object(NowOn).VertexCount) As VertexDis
    If ViewMode = 1 Then
        Object(NowOn).Vertex(Object(NowOn).VertexCount).X = Snaped((frmMain.Guide.x2 - xof) / Zoom)
        Object(NowOn).Vertex(Object(NowOn).VertexCount).Y = MidY
        Object(NowOn).Vertex(Object(NowOn).VertexCount).Z = Snaped((frmMain.Guide.y2 - yof) / Zoom)
    ElseIf ViewMode = 2 Then
        Object(NowOn).Vertex(Object(NowOn).VertexCount).X = Snaped(frmMain.Guide.x2 * Zoom - xof)
        Object(NowOn).Vertex(Object(NowOn).VertexCount).Z = MidZ
        Object(NowOn).Vertex(Object(NowOn).VertexCount).Y = Snaped(frmMain.Guide.y2 * Zoom - yof)
    Else
        Object(NowOn).Vertex(Object(NowOn).VertexCount).Z = Snaped(frmMain.Guide.x2 * Zoom - xof)
        Object(NowOn).Vertex(Object(NowOn).VertexCount).Y = Snaped(frmMain.Guide.y2 * Zoom - yof)
        Object(NowOn).Vertex(Object(NowOn).VertexCount).X = MidX
    End If
    
    Dim VNEw, HoldThis
    VNEw = Object(NowOn).VertexCount
    Object(NowOn).FaceCount = Object(NowOn).FaceCount + ThaNewFace(1)
    For M = 2 To ThaNewFace(1) + 1
        HoldThis = Object(NowOn).EdgeCount
        Object(NowOn).EdgeCount = Object(NowOn).EdgeCount + 4
        ReDim Preserve Object(NowOn).Face(Object(NowOn).EdgeCount) As Integer
        Object(NowOn).Face(HoldThis + 1) = 3
        Object(NowOn).Face(HoldThis + 2) = ThaNewFace(M)
        If M = ThaNewFace(1) + 1 Then
            Object(NowOn).Face(HoldThis + 3) = ThaNewFace(2)
        Else
            Object(NowOn).Face(HoldThis + 3) = ThaNewFace(M + 1)
        End If
        Object(NowOn).Face(HoldThis + 4) = VNEw
    Next M
    FindOutLine (NowOn)
    frmMain.DrawModel
End Sub


Public Sub ExtendFace(NowOn, Face, MidFace)
    Dim FaceON As Integer, N As Integer, M As Integer, edge As Integer, Edges As Integer, xof As Integer
    Dim NewON As Integer, NewSHAPE() As Integer, NewFace As Integer, NewFaceOn As Integer, yof As Integer
    Dim OV As Integer, Xxx As Integer, Yyy As Integer, Zzz As Integer, HoldThis As Integer, HoldThat As Integer
    xof = (frmMain.View.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.View.ScaleHeight / 2) - frmMain.UDBar
    FaceON = 1
    NewON = 1
    NewFaceOn = 1
    For N = 1 To Object(NowOn).FaceCount
        Edges = Object(NowOn).Face(FaceON)
        If N <> Face Then
            ReDim Preserve NewSHAPE(FaceON) As Integer
            NewSHAPE(NewFaceOn) = Edges
            NewFaceOn = NewFaceOn + 1
        Else
            ReDim ThaNewFace(Edges + 1) As Integer
            ThaNewFace(1) = Edges
        End If
        FaceON = FaceON + 1
        For M = 1 To Edges
            edge = Object(NowOn).Face(FaceON)
            If N <> Face Then
                ReDim Preserve NewSHAPE(NewFaceOn) As Integer
                NewSHAPE(NewFaceOn) = edge
                NewFaceOn = NewFaceOn + 1
            Else
                ThaNewFace(M + 1) = edge
            End If
            
            
            FaceON = FaceON + 1
        Next M
    Next N
    Object(NowOn).EdgeCount = NewFaceOn - 1
    Object(NowOn).FaceCount = Object(NowOn).FaceCount - 1
    ReDim Object(NowOn).Face(Object(NowOn).EdgeCount) As Integer
    For N = 1 To Object(NowOn).EdgeCount
        Object(NowOn).Face(N) = NewSHAPE(N)
    Next N
    
    
    OV = Object(NowOn).VertexCount
    Object(NowOn).VertexCount = Object(NowOn).VertexCount + ThaNewFace(1)
    ReDim Preserve Object(NowOn).Vertex(Object(NowOn).VertexCount) As VertexDis
    
    Xxx = Snaped((frmMain.Guide.x2 - frmMain.Guide.x1) / Zoom)
    Yyy = Snaped((frmMain.Guide.y2 - frmMain.Guide.y1) / Zoom)
    
    For N = 1 To ThaNewFace(1)
        If ViewMode = 1 Then
            Object(NowOn).Vertex(OV + N).X = Object(NowOn).Vertex(ThaNewFace(N + 1)).X + Xxx
            Object(NowOn).Vertex(OV + N).Y = Object(NowOn).Vertex(ThaNewFace(N + 1)).Y
            Object(NowOn).Vertex((OV + N)).Z = Object(NowOn).Vertex(ThaNewFace((N + 1))).Z + Yyy
        ElseIf ViewMode = 2 Then
            Object(NowOn).Vertex(OV + N).X = Object(NowOn).Vertex(ThaNewFace(N + 1)).X + Xxx
            Object(NowOn).Vertex(OV + N).Z = Object(NowOn).Vertex(ThaNewFace(N + 1)).Z
            Object(NowOn).Vertex((OV + N)).Y = Object(NowOn).Vertex(ThaNewFace((N + 1))).Y + Yyy
        Else
            Object(NowOn).Vertex(OV + N).Z = Object(NowOn).Vertex(ThaNewFace(N + 1)).Z + Xxx
            Object(NowOn).Vertex(OV + N).Y = Object(NowOn).Vertex(ThaNewFace(N + 1)).Y + Yyy
            Object(NowOn).Vertex((OV + N)).X = Object(NowOn).Vertex(ThaNewFace((N + 1))).X
        End If
    Next N
    
    
    
    
    
    
    For N = 2 To ThaNewFace(1) + 1
        Object(NowOn).FaceCount = Object(NowOn).FaceCount + 1
        HoldThis = Object(NowOn).EdgeCount
        Object(NowOn).EdgeCount = Object(NowOn).EdgeCount + 5
        ReDim Preserve Object(NowOn).Face(Object(NowOn).EdgeCount) As Integer
        Object(NowOn).Face(HoldThis + 1) = 4
        Object(NowOn).Face(HoldThis + 2) = ThaNewFace(N)
        If N = ThaNewFace(1) + 1 Then
            Object(NowOn).Face(HoldThis + 3) = ThaNewFace(2)
            Object(NowOn).Face(HoldThis + 4) = OV + 1
        Else
            Object(NowOn).Face(HoldThis + 3) = ThaNewFace(N + 1)
            Object(NowOn).Face(HoldThis + 4) = OV + N
        End If
        Object(NowOn).Face(HoldThis + 5) = OV + N - 1
    Next N
    
    Object(NowOn).FaceCount = Object(NowOn).FaceCount + 1
    ReDim Preserve Object(NowOn).Face((Object(NowOn).EdgeCount + ThaNewFace(1) + 1)) As Integer
    HoldThis = Object(NowOn).EdgeCount
    HoldThat = Object(NowOn).VertexCount - ThaNewFace(1)
    Object(NowOn).Face(HoldThis + 1) = ThaNewFace(1)
    
    
    For N = 2 To ThaNewFace(1) + 1
    Object(NowOn).Face(HoldThis + N) = (HoldThat + N - 1)
    Next N
    
    Object(NowOn).EdgeCount = Object(NowOn).EdgeCount + ThaNewFace(1) + 1
    
    
    FindOutLine (NowOn)
    frmMain.DrawModel
End Sub
