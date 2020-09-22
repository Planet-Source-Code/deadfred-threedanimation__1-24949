Attribute VB_Name = "Add_a"
Dim x1 As Single, x2 As Single, y1 As Single, y2 As Single, z1 As Single, z2 As Single


Public Sub Tourous()
    Add = AddObject
    GetCorners
    HFace = frmMain.ShpProp(1)
    VFace = frmMain.ShpProp(5)
    xof = (frmMain.View.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.View.ScaleHeight / 2) - frmMain.UDBar
    Object(Add).Used = True
    Object(Add).FaceCount = HFace * VFace
    Object(Add).EdgeCount = (HFace * VFace) * 5
    Object(Add).VertexCount = HFace * VFace
    Object(Add).GroupCount = 0
    Object(Add).Selected = True
    ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
    ReDim Object(Add).Vertex(Object(Add).VertexCount) As VertexDis
    frmMain.View.DrawStyle = 0
    VertexOn = 0
    HAxis = frmMain.Axis.y1 - yof
    VAxis = frmMain.Axis.x1 - xof
    CenX = (x1 + x2) / 2
    CenY = (y1 + y2) / 2
    CenZ = (z1 + z2) / 2
    Dim ReverseThis As Boolean
    ReverseThis = False
    For Eg = 0 To 359 Step 360 / HFace
        If frmMain.Axis.Tag = "X" Then
            Ega = Eg + (180 / HFace) - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye): Xx1 = Xx1 * (x1 - x2) * 0.5
            Xx1 = Xx1 + (x1 + x2) / 2: zz1 = -Cos(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * 0.5: zz1 = zz1 + (z1 + z2) / 2
            zz1 = zz1 - HAxis
            For a = 0 To 359 Step 360 / VFace
                xx2 = Cos(a / Pye) * zz1: yy2 = Sin(a / Pye) * zz1: zz2 = Xx1
                VertexOn = VertexOn + 1
                If ViewMode = 1 Then
                    Object(Add).Vertex(VertexOn).X = zz2
                    Object(Add).Vertex(VertexOn).Y = yy2
                    Object(Add).Vertex(VertexOn).Z = xx2 + HAxis
                    If CenZ < HAxis Then ReverseThis = True
                End If
            Next a
        End If
        If frmMain.Axis.Tag = "Y" Then
            Ega = Eg + (180 / HFace) - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye): Xx1 = Xx1 * (x1 - x2) * 0.5
            Xx1 = Xx1 + (x1 + x2) / 2: zz1 = -Cos(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * 0.5: zz1 = zz1 + (z1 + z2) / 2
            Xx1 = Xx1 - VAxis
            For a = 0 To 359 Step 360 / VFace
                xx2 = (Cos(a / Pye) * Xx1)
                yy2 = (Sin(a / Pye) * Xx1)
                zz2 = zz1
                VertexOn = VertexOn + 1
                If ViewMode = 1 Then
                    Object(Add).Vertex(VertexOn).X = xx2 + VAxis
                    Object(Add).Vertex(VertexOn).Y = yy2
                    Object(Add).Vertex(VertexOn).Z = zz2
                    If CenX < VAxis Then ReverseThis = True
                End If
            Next a
        End If
    Next Eg
    FaceON = 1
    For FF = 0 To VFace - 2
        For n = 1 To HFace - 1
            Mup = (n * VFace) - VFace
            Object(Add).Face(FaceON) = 4:                                    FaceON = FaceON + 1
            Object(Add).Face(FaceON) = 1 + Mup + FF:                    FaceON = FaceON + 1
            Object(Add).Face(FaceON) = 2 + Mup + FF:                    FaceON = FaceON + 1
            Object(Add).Face(FaceON) = (2 + VFace + Mup + FF):      FaceON = FaceON + 1
            Object(Add).Face(FaceON) = 1 + VFace + Mup + FF:        FaceON = FaceON + 1
        Next n
        Mup = (n * VFace) - VFace
        Object(Add).Face(FaceON) = 4:                                   FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 1 + Mup + FF:                   FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 2 + Mup + FF:                   FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 2 + FF:                            FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 1 + FF:                            FaceON = FaceON + 1
    Next FF
    For n = 1 To HFace - 1
        Mup = (n * VFace) - VFace
        Object(Add).Face(FaceON) = 4:                                  FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 1 + Mup:                         FaceON = FaceON + 1
        Object(Add).Face(FaceON) = VFace + 1 + Mup:             FaceON = FaceON + 1
        Object(Add).Face(FaceON) = (VFace * 2) + Mup:           FaceON = FaceON + 1
        Object(Add).Face(FaceON) = VFace + Mup:                   FaceON = FaceON + 1
    Next n
    Object(Add).Face(FaceON) = 4:                                        FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 1:                                        FaceON = FaceON + 1
    Object(Add).Face(FaceON) = VFace:                                 FaceON = FaceON + 1
    Object(Add).Face(FaceON) = VertexOn:                             FaceON = FaceON + 1
    Object(Add).Face(FaceON) = VertexOn + 1 - VFace:            FaceON = FaceON + 1
    If ReverseThis = True Then
        FlipFaces Add
    End If
    FindOutLine (Add)
    Object(Add).Selected = True
End Sub


Public Sub Roundoid(LineLength)
    Add = AddObject
    HFace = frmMain.ShpProp(1)
    xof = (frmMain.View.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.View.ScaleHeight / 2) - frmMain.UDBar
    Object(Add).Used = True
    Object(Add).FaceCount = HFace * (LineLength - 1)
    Object(Add).EdgeCount = ((HFace) * 5) * LineLength
    Object(Add).VertexCount = LineLength * HFace
    Object(Add).GroupCount = 0
    Object(Add).Selected = True
    ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
    ReDim Object(Add).Vertex(Object(Add).VertexCount) As VertexDis
    frmMain.View.DrawStyle = 0
    VertexOn = 0
    HAxis = frmMain.Axis.y1 - yof
    VAxis = frmMain.Axis.x1 - xof
    If ViewMode = 1 Then
        If frmMain.Axis.Tag = "X" Then
            For n = 1 To LineLength
                For a = 0 To 359 Step 360 / HFace
                   XX = StoreLine(n).X - xof
                   YY = StoreLine(n).Y - yof - HAxis
                   yy1 = Sin(a / Pye) * YY
                   yy2 = Cos(a / Pye) * YY
                   VertexOn = VertexOn + 1
                   Object(Add).Vertex(VertexOn).X = XX
                   Object(Add).Vertex(VertexOn).Y = yy1
                   Object(Add).Vertex(VertexOn).Z = yy2 + HAxis
                Next a
            Next n
        End If
        If frmMain.Axis.Tag = "Y" Then
            For n = 1 To LineLength
                For a = 0 To 359 Step 360 / HFace
                   XX = StoreLine(n).X - xof - VAxis
                   YY = StoreLine(n).Y - yof
                   Xx1 = Sin(a / Pye) * XX
                   xx2 = Cos(a / Pye) * XX
                   VertexOn = VertexOn + 1
                   Object(Add).Vertex(VertexOn).X = xx2 + VAxis
                   Object(Add).Vertex(VertexOn).Y = Xx1
                   Object(Add).Vertex(VertexOn).Z = YY
                Next a
            Next n
        End If
    End If
    If ViewMode = 2 Then
        If frmMain.Axis.Tag = "X" Then
            For n = 1 To LineLength
                For a = 0 To 359 Step 360 / HFace
                   XX = StoreLine(n).X - xof
                   YY = StoreLine(n).Y - yof - HAxis
                   yy1 = Sin(a / Pye) * YY
                   yy2 = Cos(a / Pye) * YY
                   VertexOn = VertexOn + 1
                   Object(Add).Vertex(VertexOn).X = XX
                   Object(Add).Vertex(VertexOn).Z = yy1
                   Object(Add).Vertex(VertexOn).Y = yy2 + HAxis
                Next a
            Next n
        End If
        If frmMain.Axis.Tag = "Y" Then
            For n = 1 To LineLength
                For a = 0 To 359 Step 360 / HFace
                   XX = StoreLine(n).X - xof - VAxis
                   YY = StoreLine(n).Y - yof
                   Xx1 = Sin(a / Pye) * XX
                   xx2 = Cos(a / Pye) * XX
                   VertexOn = VertexOn + 1
                   Object(Add).Vertex(VertexOn).X = xx2 + VAxis
                   Object(Add).Vertex(VertexOn).Z = Xx1
                   Object(Add).Vertex(VertexOn).Y = YY
                Next a
            Next n
        End If
    End If
    If ViewMode = 3 Then
        If frmMain.Axis.Tag = "X" Then
            For n = 1 To LineLength
                For a = 0 To 359 Step 360 / HFace
                   XX = StoreLine(n).X - xof
                   YY = StoreLine(n).Y - yof - HAxis
                   yy1 = Sin(a / Pye) * YY
                   yy2 = Cos(a / Pye) * YY
                   VertexOn = VertexOn + 1
                   Object(Add).Vertex(VertexOn).Z = XX
                   Object(Add).Vertex(VertexOn).X = yy2
                   Object(Add).Vertex(VertexOn).Y = yy1 + HAxis
                Next a
            Next n
        End If
        If frmMain.Axis.Tag = "Y" Then
            For n = 1 To LineLength
                For a = 0 To 359 Step 360 / HFace
                   XX = StoreLine(n).X - xof - VAxis
                   YY = StoreLine(n).Y - yof
                   Xx1 = Sin(a / Pye) * XX
                   xx2 = Cos(a / Pye) * XX
                   VertexOn = VertexOn + 1
                   Object(Add).Vertex(VertexOn).Z = xx2 + VAxis
                   Object(Add).Vertex(VertexOn).X = Xx1
                   Object(Add).Vertex(VertexOn).Y = YY
                Next a
            Next n
        End If
    End If
   FaceON = 1
   For FFF = 1 To LineLength - 1
        FF = (FFF - 1) * HFace
        For n = 1 To HFace - 1
            Object(Add).Face(FaceON) = 4:                          FaceON = FaceON + 1
            Object(Add).Face(FaceON) = n + HFace + FF:       FaceON = FaceON + 1
            Object(Add).Face(FaceON) = n + 1 + HFace + FF:  FaceON = FaceON + 1
            Object(Add).Face(FaceON) = n + 1 + FF:              FaceON = FaceON + 1
            Object(Add).Face(FaceON) = n + FF:                   FaceON = FaceON + 1
        Next n
        Object(Add).Face(FaceON) = 4:                          FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n + HFace + FF:        FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n + 1 + FF:               FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n + 1 - HFace + FF:   FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n + FF:                    FaceON = FaceON + 1
    Next FFF

     FindOutLine (Add)


End Sub


Public Sub Sphere()
    Add = AddObject
    GetCorners
    VFace = frmMain.ShpProp(1)
    HFace = frmMain.ShpProp(5)
    Object(Add).Used = True
    Object(Add).FaceCount = (HFace * 2) + ((VFace - 2) * HFace)
    Object(Add).EdgeCount = (2 * (HFace * 4)) + ((VFace - 2) * HFace * 5)
    Object(Add).VertexCount = 2
    Object(Add).GroupCount = 0
    Object(Add).Selected = True
    ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
    ReDim Object(Add).Vertex(Object(Add).VertexCount) As VertexDis
    CXD = 140
    VertexOn = 0
    Mf = 180 / VFace
    
    If ViewMode = 1 Then
        For n = 0 To 180 Step 180 / VFace
            If n = 0 Then
                    VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                    X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                    X = (X * 0.5) + ((x1 + x2) * 0.5)
                    Z = Cos(m / Pye) * ((z2 - z1)) * Sin(n / Pye)
                    Z = (Z * 0.5)
                    Object(Add).Vertex(VertexOn).X = X
                    Object(Add).Vertex(VertexOn).Z = -Cos(n / Pye) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)
                    Object(Add).Vertex(VertexOn).Y = Z
            ElseIf n = 180 Then
                    VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                    X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                    X = (X * 0.5) + ((x1 + x2) * 0.5)
                    Z = Cos(m / Pye) * ((z2 - z1)) * Sin(n / Pye)
                    Z = (Z * 0.5)
                    Object(Add).Vertex(VertexOn).X = X
                    Object(Add).Vertex(VertexOn).Z = -Cos(n / Pye) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)
                    Object(Add).Vertex(VertexOn).Y = Z
            Else
                For m = 0 To 359 Step 360 / HFace
                        VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                        X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                        X = (X * 0.5) + ((x1 + x2) * 0.5)
                        Z = Cos(m / Pye) * ((z2 - z1)) * Sin(n / Pye)
                        Z = (Z * 0.5)
                        Object(Add).Vertex(VertexOn).X = X
                        Object(Add).Vertex(VertexOn).Z = -Cos(n / Pye) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)
                        Object(Add).Vertex(VertexOn).Y = Z
                Next m
            End If
        Next n
    ElseIf ViewMode = 2 Then
        For n = 0 To 180 Step 180 / VFace
            If n = 0 Then
                    VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                    X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                    X = (X * 0.5) + ((x1 + x2) * 0.5)
                    Z = Cos(m / Pye) * ((z2 - z1)) * Sin(n / Pye)
                    Z = (Z * 0.5)
                    Object(Add).Vertex(VertexOn).X = X
                    Object(Add).Vertex(VertexOn).Y = -Cos(n / Pye) * ((y2 - y1) * 0.5) + ((y1 + y2) * 0.5)
                    Object(Add).Vertex(VertexOn).Z = Z
            ElseIf n = 180 Then
                    VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                    X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                    X = (X * 0.5) + ((x1 + x2) * 0.5)
                    Z = Cos(m / Pye) * ((z2 - z1)) * Sin(n / Pye)
                    Z = (Z * 0.5)
                    Object(Add).Vertex(VertexOn).X = X
                    Object(Add).Vertex(VertexOn).Y = -Cos(n / Pye) * ((y2 - y1) * 0.5) + ((y1 + y2) * 0.5)
                    Object(Add).Vertex(VertexOn).Z = Z
            Else
                For m = 0 To 359 Step 360 / HFace
                        VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                        X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                        X = (X * 0.5) + ((x1 + x2) * 0.5)
                        Z = Cos(m / Pye) * ((z2 - z1)) * Sin(n / Pye)
                        Z = (Z * 0.5)
                        Object(Add).Vertex(VertexOn).X = X
                        Object(Add).Vertex(VertexOn).Y = -Cos(n / Pye) * ((y2 - y1) * 0.5) + ((y1 + y2) * 0.5)
                        Object(Add).Vertex(VertexOn).Z = Z
                Next m
            End If
        Next n
    ElseIf ViewMode = 3 Then
        For n = 0 To 180 Step 180 / VFace
            If n = 0 Then
                    VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                    X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                    X = (X * 0.5) + ((x1 + x2) * 0.5)
                    Z = Cos(m / Pye) * ((x2 - x1)) * Sin(n / Pye)
                    Z = (X * 0.5)
                    Object(Add).Vertex(VertexOn).X = Z
                    Object(Add).Vertex(VertexOn).Z = -Cos(n / Pye) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)
                    Object(Add).Vertex(VertexOn).Y = X
            ElseIf n = 180 Then
                    VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                    X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                    X = (X * 0.5) + ((x1 + x2) * 0.5)
                    Z = Cos(m / Pye) * ((x2 - x1)) * Sin(n / Pye)
                    Z = (X * 0.5)
                    Object(Add).Vertex(VertexOn).X = Z
                    Object(Add).Vertex(VertexOn).Z = -Cos(n / Pye) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)
                    Object(Add).Vertex(VertexOn).Y = X
            Else
                For m = 0 To 359 Step 360 / HFace
                        VertexOn = VertexOn + 1: If VertexOn > UBound(Object(Add).Vertex) Then ReDim Preserve Object(Add).Vertex(VertexOn) As VertexDis
                        X = (Sin(m / Pye) * ((x2 - x1)) * Sin(n / Pye))
                        X = (X * 0.5) + ((x1 + x2) * 0.5)
                    Z = Cos(m / Pye) * ((x2 - x1)) * Sin(n / Pye)
                    Z = (X * 0.5)
                        Object(Add).Vertex(VertexOn).X = Z
                    Object(Add).Vertex(VertexOn).Z = -Cos(n / Pye) * ((z2 - z1) * 0.5) + ((z1 + z2) * 0.5)
                        Object(Add).Vertex(VertexOn).Y = X
                Next m
            End If
        Next n
    End If
    
    Object(Add).VertexCount = VertexOn
    FindOutLine (Add)
    FaceON = FaceON + 1
    Md = (HFace * 2) - HFace - HFace + 1
    For n = 1 To HFace - 1
        Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 1: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = Md + n + 1: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = Md + n: FaceON = FaceON + 1
    Next n
    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 1: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = Md + n + 1 - HFace: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = Md + n: FaceON = FaceON + 1
    Md = (HFace * (VFace - 2)) + 1
    For n = 1 To HFace - 1
        Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = Md + n: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = Md + n + 1: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = VertexOn: FaceON = FaceON + 1
    Next n
    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = Md + n: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = Md + n + 1 - HFace: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = VertexOn: FaceON = FaceON + 1
    For m = 2 To VFace - 1
        Md = (HFace * m) - HFace - HFace + 1
        For n = 1 To HFace - 1
            Object(Add).Face(FaceON) = 4: FaceON = FaceON + 1
            Object(Add).Face(FaceON) = Md + n: FaceON = FaceON + 1
            Object(Add).Face(FaceON) = Md + n + 1: FaceON = FaceON + 1
            Object(Add).Face(FaceON) = Md + n + HFace + 1: FaceON = FaceON + 1
            Object(Add).Face(FaceON) = Md + n + HFace: FaceON = FaceON + 1
        Next n
        Object(Add).Face(FaceON) = 4: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = Md + n: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = Md + n + 1 - HFace: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = Md + n + 1: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = Md + n + HFace: FaceON = FaceON + 1
    Next m
End Sub

Public Sub Cube()
    GetCorners
    Add = AddObject
    Object(Add).Used = True
    Object(Add).FaceCount = 6
    Object(Add).EdgeCount = 30
    Object(Add).VertexCount = 8
    Object(Add).GroupCount = 0
    Object(Add).Selected = True
    ReDim Object(Add).Face(30) As Integer
    ReDim Object(Add).Vertex(8) As VertexDis
    
    For n = 1 To 8
        Object(Add).Vertex(n).TargetName = ""
    Next n
    
    
    Object(Add).Vertex(1).X = x1: Object(Add).Vertex(1).Y = y1: Object(Add).Vertex(1).Z = z1
    Object(Add).Vertex(2).X = x2: Object(Add).Vertex(2).Y = y1: Object(Add).Vertex(2).Z = z1
    Object(Add).Vertex(3).X = x2: Object(Add).Vertex(3).Y = y2: Object(Add).Vertex(3).Z = z1
    Object(Add).Vertex(4).X = x1: Object(Add).Vertex(4).Y = y2: Object(Add).Vertex(4).Z = z1
    Object(Add).Vertex(5).X = x1: Object(Add).Vertex(5).Y = y1: Object(Add).Vertex(5).Z = z2
    Object(Add).Vertex(6).X = x2: Object(Add).Vertex(6).Y = y1: Object(Add).Vertex(6).Z = z2
    Object(Add).Vertex(7).X = x2: Object(Add).Vertex(7).Y = y2: Object(Add).Vertex(7).Z = z2
    Object(Add).Vertex(8).X = x1: Object(Add).Vertex(8).Y = y2: Object(Add).Vertex(8).Z = z2
    Object(Add).Face(1) = 4:  Object(Add).Face(2) = 4:  Object(Add).Face(3) = 3:  Object(Add).Face(4) = 2:  Object(Add).Face(5) = 1
    Object(Add).Face(6) = 4:  Object(Add).Face(7) = 5:  Object(Add).Face(8) = 6:  Object(Add).Face(9) = 7: Object(Add).Face(10) = 8
    Object(Add).Face(11) = 4: Object(Add).Face(12) = 1: Object(Add).Face(13) = 2: Object(Add).Face(14) = 6: Object(Add).Face(15) = 5
    Object(Add).Face(16) = 4: Object(Add).Face(17) = 3: Object(Add).Face(18) = 4: Object(Add).Face(19) = 8: Object(Add).Face(20) = 7
    Object(Add).Face(21) = 4: Object(Add).Face(22) = 5: Object(Add).Face(23) = 8: Object(Add).Face(24) = 4: Object(Add).Face(25) = 1
    Object(Add).Face(26) = 4: Object(Add).Face(27) = 2: Object(Add).Face(28) = 3: Object(Add).Face(29) = 7: Object(Add).Face(30) = 6
    
     FindOutLine (Add)
    
End Sub


Public Sub Prism()
    GetCorners
    Add = AddObject
    Edges = frmMain.ShpProp(1)
    Ends = 2 * (Edges + 1)
    Object(Add).Used = True
    Object(Add).FaceCount = Edges + 2
    Object(Add).EdgeCount = (Edges * 5) + Ends
    Object(Add).VertexCount = Edges * 2
    Object(Add).GroupCount = 0

    ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
    ReDim Object(Add).Vertex(Edges * 2) As VertexDis
    Vert = 1
    For Eg = 0 To 359 Step 360 / Edges
        If ViewMode = 1 Then
            Ega = Eg + (180 / Edges) + 0.1 - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye)
            Xx1 = Xx1 * (x1 - x2) * (frmMain.ShpProp(4) / 40)
            Xx1 = Xx1 + (x1 + x2) / 2
            zz1 = -Cos(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * (frmMain.ShpProp(4) / 40)
            zz1 = zz1 + (z1 + z2) / 2
            xx2 = Sin(Ega / Pye)
            xx2 = xx2 * (x1 - x2) * (frmMain.ShpProp(3) / 40)
            xx2 = xx2 + (x1 + x2) / 2
            zz2 = -Cos(Ega / Pye)
            zz2 = zz2 * (z1 - z2) * (frmMain.ShpProp(3) / 40)
            zz2 = zz2 + (z1 + z2) / 2
            Object(Add).Vertex(Vert).X = Xx1
            Object(Add).Vertex(Vert).Y = y2
            Object(Add).Vertex(Vert).Z = zz1
            Vert = Vert + 1
            Object(Add).Vertex(Vert).X = xx2
            Object(Add).Vertex(Vert).Y = y1
            Object(Add).Vertex(Vert).Z = zz2
            Vert = Vert + 1
        End If
        If ViewMode = 2 Then
            Ega = Eg + (180 / Edges) + 0.1 - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye)
            Xx1 = Xx1 * (x1 - x2) * (frmMain.ShpProp(3) / 40)
            Xx1 = Xx1 + (x1 + x2) / 2
            yy1 = -Cos(Ega / Pye)
            yy1 = yy1 * (y1 - y2) * (frmMain.ShpProp(3) / 40)
            yy1 = yy1 + (y1 + y2) / 2
            xx2 = Sin(Ega / Pye)
            xx2 = xx2 * (x1 - x2) * (frmMain.ShpProp(4) / 40)
            xx2 = xx2 + (x1 + x2) / 2
            yy2 = -Cos(Ega / Pye)
            yy2 = yy2 * (y1 - y2) * (frmMain.ShpProp(4) / 40)
            yy2 = yy2 + (y1 + y2) / 2
            Object(Add).Vertex(Vert).X = Xx1
            Object(Add).Vertex(Vert).Y = yy1
            Object(Add).Vertex(Vert).Z = z1
            Vert = Vert + 1
            Object(Add).Vertex(Vert).X = xx2
            Object(Add).Vertex(Vert).Y = yy2
            Object(Add).Vertex(Vert).Z = z2
            Vert = Vert + 1
        End If
        If ViewMode = 3 Then
            Ega = Eg + (180 / Edges) + 0.1 - (frmMain.ShpProp(2) * 5)
            yy1 = -Cos(Ega / Pye)
            yy1 = yy1 * (y1 - y2) * (frmMain.ShpProp(3) / 40)
            yy1 = yy1 + (y1 + y2) / 2
            zz1 = Sin(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * (frmMain.ShpProp(3) / 40)
            zz1 = zz1 + (z1 + z2) / 2
            yy2 = -Cos(Ega / Pye)
            yy2 = yy2 * (y1 - y2) * (frmMain.ShpProp(4) / 40)
            yy2 = yy2 + (y1 + y2) / 2
            zz2 = Sin(Ega / Pye)
            zz2 = zz2 * (z1 - z2) * (frmMain.ShpProp(4) / 40)
            zz2 = zz2 + (z1 + z2) / 2
            Object(Add).Vertex(Vert).X = x2
            Object(Add).Vertex(Vert).Y = yy1
            Object(Add).Vertex(Vert).Z = zz1
            Vert = Vert + 1
            Object(Add).Vertex(Vert).X = x1
            Object(Add).Vertex(Vert).Y = yy2
            Object(Add).Vertex(Vert).Z = zz2
            Vert = Vert + 1
        End If
    Next Eg
    FaceON = 1
    For n = 1 To Edges - 1
        nn = n * 2
        Object(Add).Face(FaceON) = 4: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = (nn + 1): FaceON = FaceON + 1
        Object(Add).Face(FaceON) = (nn + 2): FaceON = FaceON + 1
        Object(Add).Face(FaceON) = nn: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = (nn - 1): FaceON = FaceON + 1
    Next n
    nn = n * 2
    Object(Add).Face(FaceON) = 4: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 1: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 2: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = nn: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = nn - 1: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = Edges: FaceON = FaceON + 1
    For n = (Edges * 2) - 1 To 1 Step -2
        Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
    Next n
    Object(Add).Face(FaceON) = Edges: FaceON = FaceON + 1
    For n = 1 To (Edges * 2) - 1 Step 2
        Object(Add).Face(FaceON) = n + 1: FaceON = FaceON + 1
    Next n
    
     FindOutLine (Add)
    
    
End Sub

Public Sub Plane()
    GetCorners
    Add = AddObject
    Edges = frmMain.ShpProp(1)
    Object(Add).Used = True
    Object(Add).FaceCount = 1
    Object(Add).EdgeCount = Edges + 1
    Object(Add).VertexCount = Edges
    Object(Add).GroupCount = 0

    ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
    ReDim Object(Add).Vertex(Edges) As VertexDis
    Vert = 1: FaceON = 1
    For Eg = 0 To 359 Step 360 / Edges
        If ViewMode = 1 Then
            Ega = Eg + (180 / Edges) - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye)
            Xx1 = Xx1 * (x1 - x2) * 0.5
            Xx1 = Xx1 + (x1 + x2) / 2
            zz1 = -Cos(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * 0.5
            zz1 = zz1 + (z1 + z2) / 2
            Object(Add).Vertex(Vert).X = Xx1
            Object(Add).Vertex(Vert).Y = 0
            Object(Add).Vertex(Vert).Z = zz1
            Vert = Vert + 1
        End If
        If ViewMode = 2 Then
            Ega = Eg + (180 / Edges) - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye)
            Xx1 = Xx1 * (x1 - x2) * 0.5
            Xx1 = Xx1 + (x1 + x2) / 2
            yy1 = -Cos(Ega / Pye)
            yy1 = yy1 * (y1 - y2) * 0.5
            yy1 = yy1 + (y1 + y2) / 2
            Object(Add).Vertex(Vert).X = Xx1
            Object(Add).Vertex(Vert).Y = yy1
            Object(Add).Vertex(Vert).Z = 0
            Vert = Vert + 1
        End If
        If ViewMode = 3 Then
            Ega = Eg + (180 / Edges) + (frmMain.ShpProp(2) * 5)
            yy1 = -Cos(Ega / Pye)
            yy1 = yy1 * (y1 - y2) * 0.5
            yy1 = yy1 + (y1 + y2) / 2
            zz1 = -Sin(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * 0.5
            zz1 = zz1 + (z1 + z2) / 2
            Object(Add).Vertex(Vert).X = 0
            Object(Add).Vertex(Vert).Y = yy1
            Object(Add).Vertex(Vert).Z = zz1
            Vert = Vert + 1
        End If
    Next Eg
    Object(Add).Face(FaceON) = Edges: FaceON = FaceON + 1
    For n = 1 To Edges
        Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
    Next n
    
     FindOutLine (Add)
    
End Sub

Public Sub Pyramid()
    GetCorners
    Add = AddObject
    Edges = frmMain.ShpProp(1)
    Object(Add).Used = True
    Object(Add).FaceCount = 1 + Edges
    Object(Add).EdgeCount = 1 + Edges + (Edges * 4)
    Object(Add).VertexCount = Edges + 1
    Object(Add).GroupCount = 0

    ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
    ReDim Object(Add).Vertex(Edges + 1) As VertexDis
    Vert = 1: FaceON = 1
    If ViewMode = 1 Then
        Object(Add).Vertex(Vert).X = (x1 + x2) / 2
        Object(Add).Vertex(Vert).Y = y1
        Object(Add).Vertex(Vert).Z = (z1 + z2) / 2
        Vert = Vert + 1
        For Eg = 359 To 0 Step -(360 / Edges)
            Ega = Eg + (180 / Edges) - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye)
            Xx1 = Xx1 * (x1 - x2) * 0.5
            Xx1 = Xx1 + (x1 + x2) / 2
            zz1 = -Cos(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * 0.5
            zz1 = zz1 + (z1 + z2) / 2
            Object(Add).Vertex(Vert).X = Xx1
            Object(Add).Vertex(Vert).Y = y2
            Object(Add).Vertex(Vert).Z = zz1
            Vert = Vert + 1
        Next Eg
    End If
    If ViewMode = 2 Then
        Object(Add).Vertex(Vert).X = (x1 + x2) / 2
        Object(Add).Vertex(Vert).Y = (y1 + y2) / 2
        Object(Add).Vertex(Vert).Z = z1
        Vert = Vert + 1
        For Eg = 0 To 359 Step 360 / Edges
            Ega = Eg + (180 / Edges) - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye)
            Xx1 = Xx1 * (x1 - x2) * 0.5
            Xx1 = Xx1 + (x1 + x2) / 2
            yy1 = -Cos(Ega / Pye)
            yy1 = yy1 * (y1 - y2) * 0.5
            yy1 = yy1 + (y1 + y2) / 2
            Object(Add).Vertex(Vert).X = Xx1
            Object(Add).Vertex(Vert).Y = yy1
            Object(Add).Vertex(Vert).Z = z2
            Vert = Vert + 1
        Next Eg
    End If
    If ViewMode = 3 Then
        Object(Add).Vertex(Vert).X = x1
        Object(Add).Vertex(Vert).Y = (y1 + y2) / 2
        Object(Add).Vertex(Vert).Z = (z1 + z2) / 2
        Vert = Vert + 1
        For Eg = 0 To 359 Step 360 / Edges
            Ega = Eg + (180 / Edges) + (frmMain.ShpProp(2) * 5)
            yy1 = -Cos(Ega / Pye)
            yy1 = yy1 * (y1 - y2) * 0.5
            yy1 = yy1 + (y1 + y2) / 2
            zz1 = -Sin(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * 0.5
            zz1 = zz1 + (z1 + z2) / 2
            Object(Add).Vertex(Vert).X = x2
            Object(Add).Vertex(Vert).Y = yy1
            Object(Add).Vertex(Vert).Z = zz1
            Vert = Vert + 1
        Next Eg
    End If
    Object(Add).Face(FaceON) = Edges: FaceON = FaceON + 1
    For n = 2 To Edges + 1
        Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
    Next n
    For n = 2 To Edges
        Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n + 1: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 1: FaceON = FaceON + 1
    Next n
    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 2: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 1: FaceON = FaceON + 1
     FindOutLine (Add)
    
End Sub

Public Sub Dimond()
    GetCorners
    Add = AddObject
    Edges = frmMain.ShpProp(1)
    Object(Add).Used = True
    Object(Add).FaceCount = Edges * 2
    Object(Add).EdgeCount = (Edges * 4) * 2
    Object(Add).VertexCount = Edges + 2
    Object(Add).GroupCount = 0

    ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
    ReDim Object(Add).Vertex(Edges + 2) As VertexDis
    Vert = 1: FaceON = 1
    If ViewMode = 1 Then
        Object(Add).Vertex(Vert).X = (x1 + x2) / 2
        Object(Add).Vertex(Vert).Y = y2
        Object(Add).Vertex(Vert).Z = (z1 + z2) / 2
        Vert = Vert + 1
        Object(Add).Vertex(Vert).X = (x1 + x2) / 2
        Object(Add).Vertex(Vert).Y = y1
        Object(Add).Vertex(Vert).Z = (z1 + z2) / 2
        Vert = Vert + 1
        For Eg = 0 To 359 Step 360 / Edges
            Ega = Eg + (180 / Edges) - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye)
            Xx1 = Xx1 * (x1 - x2) * 0.5
            Xx1 = Xx1 + (x1 + x2) / 2
            zz1 = -Cos(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * 0.5
            zz1 = zz1 + (z1 + z2) / 2
            Object(Add).Vertex(Vert).X = Xx1
            Object(Add).Vertex(Vert).Y = (y1 + y2) / 2
            Object(Add).Vertex(Vert).Z = zz1
            Vert = Vert + 1
        Next Eg
    End If
    If ViewMode = 2 Then
        Object(Add).Vertex(Vert).X = (x1 + x2) / 2
        Object(Add).Vertex(Vert).Y = (y1 + y2) / 2
        Object(Add).Vertex(Vert).Z = z1
        Vert = Vert + 1
        Object(Add).Vertex(Vert).X = (x1 + x2) / 2
        Object(Add).Vertex(Vert).Y = (y1 + y2) / 2
        Object(Add).Vertex(Vert).Z = z2
        Vert = Vert + 1
        For Eg = 0 To 359 Step 360 / Edges
            Ega = Eg + (180 / Edges) - (frmMain.ShpProp(2) * 5)
            Xx1 = Sin(Ega / Pye)
            Xx1 = Xx1 * (x1 - x2) * 0.5
            Xx1 = Xx1 + (x1 + x2) / 2
            yy1 = -Cos(Ega / Pye)
            yy1 = yy1 * (y1 - y2) * 0.5
            yy1 = yy1 + (y1 + y2) / 2
            Object(Add).Vertex(Vert).X = Xx1
            Object(Add).Vertex(Vert).Y = yy1
            Object(Add).Vertex(Vert).Z = (z1 + z2) / 2
            Vert = Vert + 1
        Next Eg
    End If
    If ViewMode = 3 Then
        Object(Add).Vertex(Vert).X = x1
        Object(Add).Vertex(Vert).Y = (y1 + y2) / 2
        Object(Add).Vertex(Vert).Z = (z1 + z2) / 2
        Vert = Vert + 1
        Object(Add).Vertex(Vert).X = x2
        Object(Add).Vertex(Vert).Y = (y1 + y2) / 2
        Object(Add).Vertex(Vert).Z = (z1 + z2) / 2
        Vert = Vert + 1
        For Eg = 0 To 359 Step 360 / Edges
            Ega = Eg + (180 / Edges) + (frmMain.ShpProp(2) * 5)
            yy1 = -Cos(Ega / Pye)
            yy1 = yy1 * (y1 - y2) * 0.5
            yy1 = yy1 + (y1 + y2) / 2
            zz1 = -Sin(Ega / Pye)
            zz1 = zz1 * (z1 - z2) * 0.5
            zz1 = zz1 + (z1 + z2) / 2
            Object(Add).Vertex(Vert).X = (x1 + x2) / 2
            Object(Add).Vertex(Vert).Y = yy1
            Object(Add).Vertex(Vert).Z = zz1
            Vert = Vert + 1
        Next Eg
    End If


    For n = 3 To Edges + 1
        Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n + 1: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 1: FaceON = FaceON + 1
    Next n
    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 1: FaceON = FaceON + 1


    For n = 3 To Edges + 1
        Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = 2: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
        Object(Add).Face(FaceON) = n + 1: FaceON = FaceON + 1
    Next n
    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 2: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = n: FaceON = FaceON + 1
    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
    
     FindOutLine (Add)
End Sub

Private Sub GetCorners()
    DeselectAll
    Model.Saved = False
    xof = frmMain.View.ScaleWidth / 2: yof = frmMain.View.ScaleHeight / 2
    If frmMain.Guide.x1 < frmMain.Guide.x2 And frmMain.Guide.y1 > frmMain.Guide.y2 Then
        tmp = frmMain.Guide.y1
        frmMain.Guide.y1 = frmMain.Guide.y2
        frmMain.Guide.y2 = tmp
    End If
    If frmMain.Guide.x1 > frmMain.Guide.x2 And frmMain.Guide.y1 < frmMain.Guide.y2 Then
        tmp = frmMain.Guide.y1
        frmMain.Guide.y1 = frmMain.Guide.y2
        frmMain.Guide.y2 = tmp
    End If
    GX1 = Int(((frmMain.Guide.x1) - xof) + frmMain.LRBar) / Zoom
    GX2 = Int(((frmMain.Guide.x2) - xof) + frmMain.LRBar) / Zoom
    GY1 = Int(((frmMain.Guide.y1) - yof) + frmMain.UDBar) / Zoom
    GY2 = Int(((frmMain.Guide.y2) - yof) + frmMain.UDBar) / Zoom
    
    Object(AddObject).Selected = True
    If ViewMode = 1 Then x1 = GX1: x2 = GX2: y1 = -50: y2 = 50: z1 = GY1: z2 = GY2
    If ViewMode = 2 Then x1 = GX1: x2 = GX2: z1 = -50: z2 = 50: y1 = GY1: y2 = GY2
    If ViewMode = 3 Then x1 = -50: x2 = 50: z1 = GX1: z2 = GX2: y1 = GY1: y2 = GY2
End Sub
