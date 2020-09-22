Attribute VB_Name = "Drawings"
Option Explicit
Public StoreLine(255) As StoredLine
Public LineLength As Byte

Public Sub AlineAxis(Posit1, Posit2)
 Dim Linz, Posit
 Linz = frmMain.view.ScaleHeight
 If frmMain.view.ScaleWidth > frmMain.view.ScaleHeight Then
  Linz = frmMain.view.ScaleWidth
 End If
 If frmMain.Axis.Tag = "X" Then
  frmMain.Axis.x1 = 0: frmMain.Axis.X2 = frmMain.view.ScaleWidth
  If Posit2 <> vbNull Then frmMain.Axis.y1 = Posit2: frmMain.Axis.Y2 = Posit2
 ElseIf frmMain.Axis.Tag = "XY" Then
  frmMain.Axis.y1 = 0 + (Posit2 - (frmMain.view.ScaleHeight / 2))
  frmMain.Axis.Y2 = Linz + (Posit2 - (frmMain.view.ScaleHeight / 2))
  frmMain.Axis.x1 = 0
  frmMain.Axis.X2 = Linz
 ElseIf frmMain.Axis.Tag = "Y" Then
  If Posit <> vbNull Then frmMain.Axis.x1 = Posit1: frmMain.Axis.X2 = Posit1
  frmMain.Axis.y1 = 0: frmMain.Axis.Y2 = Linz
 Else
  frmMain.Axis.y1 = 0 + (Posit2 - (frmMain.view.ScaleHeight / 2))
  frmMain.Axis.Y2 = Linz + (Posit2 - (frmMain.view.ScaleHeight / 2))
  frmMain.Axis.X2 = 0
  frmMain.Axis.x1 = Linz
 End If
 'frmMain.Axis.BorderColor = Colours(6)
End Sub

Public Sub DrawGuideLines(X, Y)
    Dim xof As Integer, yof As Integer, NowOn As Integer
    xof = (frmMain.view.ScaleWidth / 2) - frmMain.LRBar - X
    yof = (frmMain.view.ScaleHeight / 2) - frmMain.UDBar - Y
    
  frmMain.view.Line (Model.Min.X * Zoom + xof, Model.Min.Y * Zoom + yof)-(Model.Max.X * Zoom + xof, Model.Max.Y * Zoom + yof), Colours(3), B
    
    
    If ViewMode = 1 Then
        For NowOn = 1 To cstTotalObjects
            If Object(NowOn).Selected = True And Object(NowOn).Used = True Then
                frmMain.view.Line (Object(NowOn).Min.X * Zoom + xof, Object(NowOn).Min.Z * Zoom + yof)-(Object(NowOn).Max.X * Zoom + xof, Object(NowOn).Max.Z * Zoom + yof), , B
            End If
        Next NowOn
        For NowOn = 1 To cstTotalJoints
            If BaseFrame(NowOn).Selected = True And Object(NowOn).Used = True Then
                frmMain.view.Line (BaseFrame(NowOn).Position.X * Zoom + xof - 5, BaseFrame(NowOn).Position.Z * Zoom + yof - 5)-(BaseFrame(NowOn).Position.X * Zoom + xof + 5, BaseFrame(NowOn).Position.Z * Zoom + yof + 5), , B
                frmMain.view.Circle (BaseFrame(NowOn).Position.X * Zoom + xof, BaseFrame(NowOn).Position.Z * Zoom + yof), 4
            End If
        Next NowOn
    ElseIf ViewMode = 2 Then
        For NowOn = 1 To cstTotalObjects
            If Object(NowOn).Selected = True And Object(NowOn).Used = True Then
                frmMain.view.Line (Object(NowOn).Min.X * Zoom + xof, Object(NowOn).Min.Y * Zoom + yof)-(Object(NowOn).Max.X * Zoom + xof, Object(NowOn).Max.Y * Zoom + yof), , B
            End If
        Next NowOn
        For NowOn = 1 To cstTotalJoints
            If BaseFrame(NowOn).Selected = True And Object(NowOn).Used = True Then
                'frmMain.View.Line (BaseFrame(NowOn).Position.X * Zoom + xof - 5, BaseFrame(NowOn).Position.Z * Zoom + yof - 5)-(BaseFrame(NowOn).Position.X * Zoom + xof + 5, BaseFrame(NowOn).Position.Z * Zoom + yof + 5), , B
                frmMain.view.Circle (BaseFrame(NowOn).Position.X * Zoom + xof, BaseFrame(NowOn).Position.Y * Zoom + yof), 4
            End If
        Next NowOn
    Else
        For NowOn = 1 To cstTotalObjects
            If Object(NowOn).Selected = True Then
                frmMain.view.Line (Object(NowOn).Min.Z * Zoom + xof, Object(NowOn).Min.Y * Zoom + yof)-(Object(NowOn).Max.Z * Zoom + xof, Object(NowOn).Max.Y * Zoom + yof), , B
            End If
        Next NowOn
        For NowOn = 1 To cstTotalJoints
            If BaseFrame(NowOn).Selected = True Then
                'frmMain.View.Line (BaseFrame(NowOn).Position.X * Zoom + xof - 5, BaseFrame(NowOn).Position.Z * Zoom + yof - 5)-(BaseFrame(NowOn).Position.X * Zoom + xof + 5, BaseFrame(NowOn).Position.Z * Zoom + yof + 5), , B
                frmMain.view.Circle (BaseFrame(NowOn).Position.Z * Zoom + xof, BaseFrame(NowOn).Position.Y * Zoom + yof), 4
            End If
        Next NowOn
    End If
End Sub

Public Sub DrawGuide()
 Dim N As Integer, x1 As Integer, X2 As Integer, y1 As Integer, Y2 As Integer, nX1 As Integer, nX2 As Integer
 Dim nY1 As Integer, nY2 As Integer, X As Integer, Y As Integer, Xw As Integer, Yw As Integer, Py As Integer
 Dim totangle As Single, Ang As Single, NextRow As Integer, ReduceMe1 As Single, ReduceMe2 As Single
 Dim x3 As Integer, X4 As Integer, Y3 As Integer, Y4 As Integer, MidCirc As Integer
 frmMain.view.AutoRedraw = False
 frmMain.view.Cls
 frmMain.view.DrawStyle = 2
 If frmMain.ShapeList.Text = "Roundoid" Then
  frmMain.view.AutoRedraw = True
  For N = 1 To LineLength - 1
    x1 = StoreLine(N).X
    y1 = StoreLine(N).Y
    X2 = StoreLine(N + 1).X
    Y2 = StoreLine(N + 1).Y
    If frmMain.Axis.Tag = "Y" Then
     nX1 = x1 - ((x1 - frmMain.Axis.x1) * 2)
     nX2 = X2 - ((X2 - frmMain.Axis.x1) * 2)
     nY1 = y1
     nY2 = Y2
    ElseIf frmMain.Axis.Tag = "X" Then
     nX1 = x1
     nX2 = X2
     nY1 = y1 - ((y1 - frmMain.Axis.y1) * 2)
     nY2 = Y2 - ((Y2 - frmMain.Axis.y1) * 2)
    ElseIf frmMain.Axis.Tag = "XY" Then
    Else
     End If
     frmMain.view.Line (x1, y1)-(X2, Y2), Colours(3)
     frmMain.view.Line (nX1, nY1)-(nX2, nY2), Colours(5)
     frmMain.view.Line (nX1, nY1)-(x1, y1), Colours(5)
     frmMain.view.Line (nX2, nY2)-(X2, Y2), Colours(5)
    Next N
    frmMain.view.AutoRedraw = False
    Else
     If frmMain.Guide.Tag = "False" Then Exit Sub
     If frmMain.ShapeList.Text = "Cube" Then
      frmMain.view.Line (frmMain.Guide.x1 - 1, frmMain.Guide.y1 - 1)-(frmMain.Guide.X2 + 1, frmMain.Guide.Y2 + 1), Colours(4), B
      frmMain.view.Line (frmMain.Guide.x1, frmMain.Guide.y1)-(frmMain.Guide.X2, frmMain.Guide.Y2), Colours(3), B
     Else
      If frmMain.ShapeList.Text <> "Torous" Then frmMain.view.Line (frmMain.Guide.x1, frmMain.Guide.y1)-(frmMain.Guide.X2, frmMain.Guide.Y2), Colours(3), B
      X = (frmMain.Guide.x1 + frmMain.Guide.X2) / 2
      Y = (frmMain.Guide.y1 + frmMain.Guide.Y2) / 2
      Xw = (frmMain.Guide.X2 - frmMain.Guide.x1) / 2
      Yw = (frmMain.Guide.Y2 - frmMain.Guide.y1) / 2
      Py = (22 / 7) * 18.2
      totangle = 360
      If frmMain.ShapeList.Text = "Torous" Then
       Ang = 180 / frmMain.ShpProp(1).Value + (frmMain.ShpProp(2).Value * 5)
       For N = 0 To 359 Step 360 / frmMain.ShpProp(1).Value
        NextRow = 360 / frmMain.ShpProp(1).Value
        x1 = (Sin((N + Ang) / Py) * Xw) + X
        y1 = (Cos((N + Ang) / Py) * Yw) + Y
        X2 = (Sin((N + Ang + NextRow) / Py) * Xw) + X
        Y2 = (Cos((N + Ang + NextRow) / Py) * Yw) + Y
        If frmMain.Axis.Tag = "Y" Then
         nX1 = x1 - ((x1 - frmMain.Axis.x1) * 2)
         nX2 = X2 - ((X2 - frmMain.Axis.x1) * 2)
         nY1 = y1
         nY2 = Y2
        ElseIf frmMain.Axis.Tag = "X" Then
         nX1 = x1
         nX2 = X2
         nY1 = y1 - ((y1 - frmMain.Axis.y1) * 2)
         nY2 = Y2 - ((Y2 - frmMain.Axis.y1) * 2)
        End If
        frmMain.view.Line (nX1, nY1)-(nX2, nY2), Colours(5)
        frmMain.view.Line (x1, y1)-(nX1, nY1), Colours(5)
        frmMain.view.Line (x1, y1)-(X2, Y2), Colours(4)
       Next N
       frmMain.view.Line (frmMain.Guide.x1, frmMain.Guide.y1)-(frmMain.Guide.X2, frmMain.Guide.Y2), Colours(3), B
      Exit Sub
     ElseIf frmMain.ShapeList.Text = "Prism" Then
      Ang = 180 / frmMain.ShpProp(1).Value + (frmMain.ShpProp(2).Value * 5)
      For N = 0 To 359 Step 360 / frmMain.ShpProp(1).Value
       NextRow = 360 / frmMain.ShpProp(1).Value
       ReduceMe1 = 100 / (frmMain.ShpProp(3) * 5)
       ReduceMe2 = 100 / (frmMain.ShpProp(4) * 5)
       x1 = (Sin((N + Ang) / Py) * Xw) / ReduceMe1
       y1 = (Cos((N + Ang) / Py) * Yw) / ReduceMe1
       X2 = (Sin((N + Ang + NextRow) / Py) * Xw) / ReduceMe1
       Y2 = (Cos((N + Ang + NextRow) / Py) * Yw) / ReduceMe1
       x3 = (Sin((N + Ang) / Py) * Xw) / ReduceMe2
       Y3 = (Cos((N + Ang) / Py) * Yw) / ReduceMe2
       X4 = (Sin((N + Ang + NextRow) / Py) * Xw) / ReduceMe2
       Y4 = (Cos((N + Ang + NextRow) / Py) * Yw) / ReduceMe2
       frmMain.view.Line (x1 + X, y1 + Y)-(X2 + X, Y2 + Y), Colours(4)
       frmMain.view.Line (x1 + X, y1 + Y)-(x3 + X, Y3 + Y), Colours(4)
       frmMain.view.Line (x3 + X, Y3 + Y)-(X4 + X, Y4 + Y), Colours(4)
      Next N
     ElseIf frmMain.ShapeList.Text = "Dimond" Or frmMain.ShapeList.Text = "Pyramid" Or frmMain.ShapeList.Text = "Plane" Then
      Ang = 180 / frmMain.ShpProp(1).Value + (frmMain.ShpProp(2).Value * 5)
      For N = 0 To 359 Step 360 / frmMain.ShpProp(1).Value
       NextRow = 360 / frmMain.ShpProp(1).Value
       x1 = (Sin((N + Ang) / Py) * Xw)
       y1 = (Cos((N + Ang) / Py) * Yw)
       X2 = (Sin((N + Ang + NextRow) / Py) * Xw)
       Y2 = (Cos((N + Ang + NextRow) / Py) * Yw)
       frmMain.view.Line (x1 + X, y1 + Y)-(X2 + X, Y2 + Y), Colours(4)
       If frmMain.ShapeList.Text = "Pyramid" Or frmMain.ShapeList.Text = "Dimond" Then
        frmMain.view.Line (x1 + X, y1 + Y)-(X, Y), Colours(4)
       End If
      Next N
     Else
      Ang = -(180 / frmMain.ShpProp(1).Value + (frmMain.ShpProp(2).Value * 5))
      NextRow = 180 / frmMain.ShpProp(1).Value
      For N = NextRow To 180 Step 180 / frmMain.ShpProp(1).Value
       x1 = (Sin((N + Ang) / Py) * Xw)
       y1 = (Cos((N + Ang) / Py) * Yw)
       X2 = (Sin((N + Ang + NextRow) / Py) * Xw)
       Y2 = (Cos((N + Ang + NextRow) / Py) * Yw)
       If N = NextRow Then
        MidCirc = (Sin((NextRow + Ang) / Py) * Xw)
       End If
       nX1 = x1 - ((x1 - MidCirc) * 2)
       nY1 = y1
       nX2 = X2 - ((X2 - MidCirc) * 2)
       nY2 = Y2
       frmMain.view.Line (nX1 + X, nY1 + Y)-(nX2 + X, nY2 + Y), Colours(5)
       frmMain.view.Line (nX1 + X, nY1 + Y)-(x1 + X, y1 + Y), Colours(5)
       frmMain.view.Line (x1 + X, y1 + Y)-(X2 + X, Y2 + Y), Colours(4)
      Next N
     End If
    End If
   End If
 End Sub
 
Public Sub DrawFromTop()
    Dim xof As Integer, yof As Integer, NowOn As Integer, N As Integer, M As Integer, Cowla As Double
    Dim FirstSelect As Boolean, ShowPoints As Boolean, edge As Integer, EdgeCount As Integer, StartOfFace As Integer
    Dim Corner1 As Integer, Corner2 As Integer, x1 As Integer, y1 As Integer, X2 As Integer, Y2 As Integer
    Dim HighLightCenter As Boolean, HighlightCorner As Boolean, FillCorner As Boolean, FaceON As Integer
    Dim CenX As Integer, CenY As Integer, CenZ As Integer, X As Integer, Y As Integer, Z As Integer
    Dim Secound As Integer, Tagr As Integer
    
    FindObjectOutline
    xof = (frmMain.view.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.view.ScaleHeight / 2) - frmMain.UDBar
    
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True Then
            Cowla = Object(NowOn).Colour
            If frmMain.opview(3).Checked = True Then Cowla = frmSettings.lblgx.BackColor
            edge = 1
            For N = 1 To Object(NowOn).FaceCount
                EdgeCount = Object(NowOn).Face(edge): edge = edge + 1
                StartOfFace = edge
                For M = 1 To EdgeCount
                    If M = EdgeCount Then
                        Corner1 = Object(NowOn).Face(edge)
                        Corner2 = Object(NowOn).Face(StartOfFace): edge = edge + 1
                        x1 = Object(NowOn).Vertex(Corner1).X * Zoom + xof
                        X2 = Object(NowOn).Vertex(Corner2).X * Zoom + xof
                        y1 = Object(NowOn).Vertex(Corner1).Z * Zoom + yof
                        Y2 = Object(NowOn).Vertex(Corner2).Z * Zoom + yof
                    Else
                        Corner1 = Object(NowOn).Face(edge): edge = edge + 1
                        Corner2 = Object(NowOn).Face(edge)
                        
                        
                        x1 = Object(NowOn).Vertex(Corner1).X * Zoom + xof
                        X2 = Object(NowOn).Vertex(Corner2).X * Zoom + xof
                        y1 = Object(NowOn).Vertex(Corner1).Z * Zoom + yof
                        Y2 = Object(NowOn).Vertex(Corner2).Z * Zoom + yof
                    End If
                    frmMain.view.DrawStyle = 0

                        frmMain.view.Line (x1, y1)-(X2, Y2), Cowla

                Next M
            Next N
            
            If Object(NowOn).Selected = True Then
                frmMain.view.DrawStyle = 2
                frmMain.view.Line (Object(NowOn).Min.X * Zoom + xof - 2, Object(NowOn).Min.Z * Zoom + yof - 2)- _
                 (Object(NowOn).Max.X * Zoom + xof + 2, Object(NowOn).Max.Z * Zoom + yof + 2), Colours(6), B
                frmMain.view.Line (Object(NowOn).Min.X * Zoom + xof - 3, Object(NowOn).Min.Z * Zoom + yof - 3)- _
                 (Object(NowOn).Max.X * Zoom + xof + 3, Object(NowOn).Max.Z * Zoom + yof + 3), Colours(6), B
                frmMain.view.DrawStyle = 0
            
            
                HighLightCenter = False
                HighlightCorner = False
                FillCorner = False
        
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(2) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(3) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(4) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(5) = True Then HighLightCenter = True
                
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(1) = True Then HighlightCorner = True: FillCorner = True
                If frmMain.chkVert = 1 Then HighlightCorner = True
    
                    If HighLightCenter = True Then
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
                            frmMain.view.Circle (CenX * Zoom + xof, CenZ * Zoom + yof), 1
                            frmMain.view.Circle (CenX * Zoom + xof, CenZ * Zoom + yof), 2
                            frmMain.view.Circle (CenX * Zoom + xof, CenZ * Zoom + yof), 3
                        Next N
                    End If
                
                    FirstSelect = True
                
                    If HighlightCorner = True Then
                        For N = 1 To Object(NowOn).VertexCount
                            X = Object(NowOn).Vertex(N).X * Zoom + xof
                            Y = Object(NowOn).Vertex(N).Z * Zoom + yof
                            frmMain.view.Circle (X, Y), 3
                            If (frmMain.chkVert = 1 And Object(NowOn).Vertex(N).Selected = True) Or FillCorner = True Then
                                frmMain.view.Circle (X, Y), 2
                                frmMain.view.Circle (X, Y), 1
                            End If
                        Next N
                    End If
            End If
        End If
    Next NowOn
    frmMain.view.DrawStyle = 0
    
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True Then
            If BaseFrame(N).Selected = True Then
                Secound = Secound + 1: If Secound = 2 Then FirstSelect = True
            End If
            X = BaseFrame(N).Position.X * Zoom + xof
            Y = BaseFrame(N).Position.Z * Zoom + yof
            frmMain.view.Circle (X, Y), 4
            If BaseFrame(N).Selected = True Then
                frmMain.view.Circle (X, Y), 3
            End If
            Tagr = FindTarget(BaseFrame(N).Target)
            If Tagr <> 0 Then
                x1 = BaseFrame(Tagr).Position.X * Zoom + xof
                y1 = BaseFrame(Tagr).Position.Z * Zoom + yof
                frmMain.view.Line (X, Y)-(x1, y1)
            End If
        End If
    Next N
    
    If FirstSelect = True Then
        frmMain.view.Line (Model.Min.X * Zoom + xof, Model.Min.Y * Zoom + yof)-(Model.Max.X * Zoom + xof, Model.Max.Y * Zoom + yof), Colours(3), B
        If frmSettings.chkHilighBox.Value = 1 Then
            x1 = Model.Min.X * Zoom + xof
            X2 = Model.Max.X * Zoom + xof
            y1 = Model.Min.Y * Zoom + yof
            Y2 = Model.Max.Y * Zoom + yof
            If (frmMain.CurrentSideBar = 6 And frmMain.edMode(1) = True) Or (frmMain.CurrentSideBar = 5 And frmMain.chkVert = 1) Then
            Else
                frmMain.view.Circle (x1, y1), 3
                frmMain.view.Circle (X2, Y2), 3
                frmMain.view.Circle (x1, Y2), 3
                frmMain.view.Circle (X2, y1), 3
                frmMain.view.Circle (x1, (y1 + Y2) * 0.5), 3
                frmMain.view.Circle (X2, (y1 + Y2) * 0.5), 3
                frmMain.view.Circle ((x1 + X2) * 0.5, Y2), 3
                frmMain.view.Circle ((x1 + X2) * 0.5, y1), 3
            End If
        End If
    End If
End Sub
 
  
Public Sub DrawFromSide()
    Dim xof As Integer, yof As Integer, NowOn As Integer, N As Integer, M As Integer, Cowla As Double
    Dim FirstSelect As Boolean, ShowPoints As Boolean, edge As Integer, EdgeCount As Integer, StartOfFace As Integer
    Dim Corner1 As Integer, Corner2 As Integer, x1 As Integer, y1 As Integer, X2 As Integer, Y2 As Integer
    Dim HighLightCenter As Boolean, HighlightCorner As Boolean, FillCorner As Boolean, FaceON As Integer
    Dim CenX As Integer, CenY As Integer, CenZ As Integer, X As Integer, Y As Integer, Z As Integer
    Dim Secound As Integer, Tagr As Integer
    
    FindObjectOutline

    xof = (frmMain.view.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.view.ScaleHeight / 2) - frmMain.UDBar
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True Then
            Cowla = Object(NowOn).Colour
            If frmMain.opview(3).Checked = True Then Cowla = frmSettings.lblgx.BackColor
            
            edge = 1
            For N = 1 To Object(NowOn).FaceCount
                EdgeCount = Object(NowOn).Face(edge): edge = edge + 1
                StartOfFace = edge
                For M = 1 To EdgeCount
                    If M = EdgeCount Then
                        Corner1 = Object(NowOn).Face(edge)
                        Corner2 = Object(NowOn).Face(StartOfFace): edge = edge + 1
                        x1 = Object(NowOn).Vertex(Corner1).X * Zoom + xof
                        X2 = Object(NowOn).Vertex(Corner2).X * Zoom + xof
                        y1 = Object(NowOn).Vertex(Corner1).Y * Zoom + yof
                        Y2 = Object(NowOn).Vertex(Corner2).Y * Zoom + yof
                    Else
                        Corner1 = Object(NowOn).Face(edge): edge = edge + 1
                        Corner2 = Object(NowOn).Face(edge)
                        x1 = Object(NowOn).Vertex(Corner1).X * Zoom + xof
                        X2 = Object(NowOn).Vertex(Corner2).X * Zoom + xof
                        y1 = Object(NowOn).Vertex(Corner1).Y * Zoom + yof
                        Y2 = Object(NowOn).Vertex(Corner2).Y * Zoom + yof
                    End If
                    frmMain.view.DrawStyle = 0
                    frmMain.view.Line (x1, y1)-(X2, Y2), Cowla
                Next M
            Next N
            
            If Object(NowOn).Selected = True Then
                FirstSelect = True
                frmMain.view.DrawStyle = 2
                frmMain.view.Line (Object(NowOn).Min.X * Zoom + xof - 2, Object(NowOn).Min.Y * Zoom + yof - 2)- _
                 (Object(NowOn).Max.X * Zoom + xof + 2, Object(NowOn).Max.Y * Zoom + yof + 2), Colours(6), B
                frmMain.view.Line (Object(NowOn).Min.X * Zoom + xof - 3, Object(NowOn).Min.Y * Zoom + yof - 3)- _
                 (Object(NowOn).Max.X * Zoom + xof + 3, Object(NowOn).Max.Y * Zoom + yof + 3), Colours(6), B
                frmMain.view.DrawStyle = 0
                
                HighLightCenter = False
                HighlightCorner = False
                FillCorner = False
                
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(2) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(3) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(4) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(5) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(1) = True Then HighlightCorner = True
                If frmMain.chkVert = 1 Then HighlightCorner = True
                If HighLightCenter = True Then
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
                        frmMain.view.Circle (CenX * Zoom + xof, CenY * Zoom + yof), 3
                        frmMain.view.Circle (CenX * Zoom + xof, CenY * Zoom + yof), 2
                        frmMain.view.Circle (CenX * Zoom + xof, CenY * Zoom + yof), 1
                    Next N
                End If
                 If HighlightCorner = True Then
                    For N = 1 To Object(NowOn).VertexCount
                        X = Object(NowOn).Vertex(N).X * Zoom + xof
                        Y = Object(NowOn).Vertex(N).Y * Zoom + yof
                        frmMain.view.Circle (X, Y), 3
                        If frmMain.edMode(1) = True Or Object(NowOn).Vertex(N).Selected = True Then
                            frmMain.view.Circle (X, Y), 2
                            frmMain.view.Circle (X, Y), 1
                        End If
                    Next N
                End If
            End If
        End If
    Next NowOn
    frmMain.view.DrawStyle = 0
     For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True Then
            If BaseFrame(N).Selected = True Then
                Secound = Secound + 1: If Secound = 2 Then FirstSelect = True
            End If
            X = BaseFrame(N).Position.X * Zoom + xof
            Y = BaseFrame(N).Position.Y * Zoom + yof
            frmMain.view.Circle (X, Y), 4
            If BaseFrame(N).Selected = True Then
                frmMain.view.Circle (X, Y), 3
            End If
            Tagr = FindTarget(BaseFrame(N).Target)
            If Tagr <> 0 Then
                x1 = BaseFrame(Tagr).Position.X * Zoom + xof
                y1 = BaseFrame(Tagr).Position.Y * Zoom + yof
                frmMain.view.Line (X, Y)-(x1, y1)
            End If
        End If
    Next N
    If FirstSelect = True Then
        frmMain.view.Line (Model.Min.X * Zoom + xof, Model.Min.Y * Zoom + yof)-(Model.Max.X * Zoom + xof, Model.Max.Y * Zoom + yof), Colours(3), B
        If frmSettings.chkHilighBox.Value = 1 Then
            x1 = Model.Min.X * Zoom + xof
            X2 = Model.Max.X * Zoom + xof
            y1 = Model.Min.Y * Zoom + yof
            Y2 = Model.Max.Y * Zoom + yof
            If (frmMain.CurrentSideBar = 6 And frmMain.edMode(1) = True) Or (frmMain.CurrentSideBar = 5 And frmMain.chkVert = 1) Then
            Else
                frmMain.view.Circle (x1, y1), 3
                frmMain.view.Circle (X2, Y2), 3
                frmMain.view.Circle (x1, Y2), 3
                frmMain.view.Circle (X2, y1), 3
                frmMain.view.Circle (x1, (y1 + Y2) * 0.5), 3
                frmMain.view.Circle (X2, (y1 + Y2) * 0.5), 3
                frmMain.view.Circle ((x1 + X2) * 0.5, Y2), 3
                frmMain.view.Circle ((x1 + X2) * 0.5, y1), 3
            End If
        End If
    End If
End Sub
 
  
Public Sub DrawRulers()
    Dim N As Integer, StpSize, Ct, Mx
    frmMain.SideRule.Cls: frmMain.TopRule.Cls
    vofs = -frmMain.UDBar + (frmMain.view.ScaleHeight / 2)
    hofs = -frmMain.LRBar + (frmMain.view.ScaleWidth / 2)
    frmMain.view.Line (hofs - 1, 0)-(hofs - 1, frmMain.view.ScaleHeight), RGB(200, 200, 200)
    frmMain.view.Line (0, vofs)-(frmMain.view.ScaleWidth, vofs), RGB(200, 200, 200)
    For N = hofs - 20 To frmMain.TopRule.ScaleWidth - 20 Step 20 * Zoom
        frmMain.TopRule.Line (N + 2, 10)-(N + 2, 20)
    Next N
    For N = hofs - 20 To 0 - 20 Step -20 * Zoom
        frmMain.TopRule.Line (N + 2, 10)-(N + 2, 20)
    Next N
    For N = 3 + vofs To frmMain.SideRule.ScaleHeight Step 20 * Zoom
        frmMain.SideRule.Line (10, N)-(20, N)
    Next N
    For N = 3 + vofs To 0 Step -20 * Zoom
        frmMain.SideRule.Line (10, N)-(20, N)
    Next N
End Sub
  
 
Public Sub DrawFromFront()
    Dim xof As Integer, yof As Integer, NowOn As Integer, N As Integer, M As Integer, Cowla As Double
    Dim FirstSelect As Boolean, ShowPoints As Boolean, edge As Integer, EdgeCount As Integer, StartOfFace As Integer
    Dim Corner1 As Integer, Corner2 As Integer, x1 As Integer, y1 As Integer, X2 As Integer, Y2 As Integer
    Dim HighLightCenter As Boolean, HighlightCorner As Boolean, FillCorner As Boolean, FaceON As Integer
    Dim CenX As Integer, CenY As Integer, CenZ As Integer, X As Integer, Y As Integer, Z As Integer
    Dim Secound As Integer, Tagr As Integer
    
    FindObjectOutline

    xof = (frmMain.view.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.view.ScaleHeight / 2) - frmMain.UDBar
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True Then
            Cowla = Object(NowOn).Colour
            If frmMain.opview(3).Checked = True Then Cowla = frmSettings.lblgx.BackColor
            edge = 1
            For N = 1 To Object(NowOn).FaceCount
                EdgeCount = Object(NowOn).Face(edge): edge = edge + 1
                StartOfFace = edge
                For M = 1 To EdgeCount
                    If M = EdgeCount Then
                        Corner1 = Object(NowOn).Face(edge)
                        Corner2 = Object(NowOn).Face(StartOfFace): edge = edge + 1
                        x1 = Object(NowOn).Vertex(Corner1).Z * Zoom + xof
                        X2 = Object(NowOn).Vertex(Corner2).Z * Zoom + xof
                        y1 = Object(NowOn).Vertex(Corner1).Y * Zoom + yof
                        Y2 = Object(NowOn).Vertex(Corner2).Y * Zoom + yof
                    Else
                        Corner1 = Object(NowOn).Face(edge): edge = edge + 1
                        Corner2 = Object(NowOn).Face(edge)
                        x1 = Object(NowOn).Vertex(Corner1).Z * Zoom + xof
                        X2 = Object(NowOn).Vertex(Corner2).Z * Zoom + xof
                        y1 = Object(NowOn).Vertex(Corner1).Y * Zoom + yof
                        Y2 = Object(NowOn).Vertex(Corner2).Y * Zoom + yof
                    End If
                    frmMain.view.DrawStyle = 0
                    frmMain.view.Line (x1, y1)-(X2, Y2), Cowla
                Next M
            Next N
            
            If Object(NowOn).Selected = True Then
            
                FirstSelect = True
                frmMain.view.DrawStyle = 2
                frmMain.view.Line (Object(NowOn).Min.Z * Zoom + xof - 2, Object(NowOn).Min.Y * Zoom + yof - 2)- _
                 (Object(NowOn).Max.Z * Zoom + xof + 2, Object(NowOn).Max.Y * Zoom + yof + 2), Colours(6), B
                frmMain.view.Line (Object(NowOn).Min.Z * Zoom + xof - 3, Object(NowOn).Min.Y * Zoom + yof - 3)- _
                 (Object(NowOn).Max.Z * Zoom + xof + 3, Object(NowOn).Max.Y * Zoom + yof + 3), Colours(6), B
                frmMain.view.DrawStyle = 0
            
                HighLightCenter = False
                HighlightCorner = False
                FillCorner = False
        
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(2) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(3) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(4) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(5) = True Then HighLightCenter = True
                If frmMain.CurrentSideBar = 6 And frmMain.edMode(1) = True Then HighlightCorner = True
                If frmMain.chkVert = 1 Then HighlightCorner = True
            
            
                If HighLightCenter = True Then
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
                        frmMain.view.Circle (CenZ * Zoom + xof, CenY * Zoom + yof), 1
                        frmMain.view.Circle (CenZ * Zoom + xof, CenY * Zoom + yof), 2
                        frmMain.view.Circle (CenZ * Zoom + xof, CenY * Zoom + yof), 3
                    Next N
                End If
            
            
                 If HighlightCorner = True Then
                    For N = 1 To Object(NowOn).VertexCount
                        X = Object(NowOn).Vertex(N).Z * Zoom + xof
                        Y = Object(NowOn).Vertex(N).Y * Zoom + yof
                        frmMain.view.Circle (X, Y), 3
                        If frmMain.edMode(1) = True Or Object(NowOn).Vertex(N).Selected = True Then
                            frmMain.view.Circle (X, Y), 2
                            frmMain.view.Circle (X, Y), 1
                        End If
                    Next N
                End If
            End If
        End If
    Next NowOn
    frmMain.view.DrawStyle = 0
    
    
'    If frmMain.opview(8).Checked = True Then
        For N = 1 To cstTotalJoints
            If BaseFrame(N).Used = True Then
                If BaseFrame(N).Selected = True Then
                    Secound = Secound + 1: If Secound = 2 Then FirstSelect = True
                End If
                X = BaseFrame(N).Position.Z * Zoom + xof
                Y = BaseFrame(N).Position.Y * Zoom + yof
                frmMain.view.Circle (X, Y), 4
                If BaseFrame(N).Selected = True Then
                    frmMain.view.Circle (X, Y), 3
                End If
                Tagr = FindTarget(BaseFrame(N).Target)
                If Tagr <> 0 Then
                    x1 = BaseFrame(Tagr).Position.Z * Zoom + xof
                    y1 = BaseFrame(Tagr).Position.Y * Zoom + yof
                    frmMain.view.Line (X, Y)-(x1, y1)
                End If
            End If
        Next N
  '  End If
    
    If FirstSelect = True Then
        frmMain.view.Line (Model.Min.X * Zoom + xof, Model.Min.Y * Zoom + yof)-(Model.Max.X * Zoom + xof, Model.Max.Y * Zoom + yof), Colours(3), B
        If frmSettings.chkHilighBox.Value = 1 Then
            x1 = Model.Min.X * Zoom + xof
            X2 = Model.Max.X * Zoom + xof
            y1 = Model.Min.Y * Zoom + yof
            Y2 = Model.Max.Y * Zoom + yof
            If (frmMain.CurrentSideBar = 6 And frmMain.edMode(1) = True) Or (frmMain.CurrentSideBar = 5 And frmMain.chkVert = 1) Then
            Else
                frmMain.view.Circle (x1, y1), 3
                frmMain.view.Circle (X2, Y2), 3
                frmMain.view.Circle (x1, Y2), 3
                frmMain.view.Circle (X2, y1), 3
                frmMain.view.Circle (x1, (y1 + Y2) * 0.5), 3
                frmMain.view.Circle (X2, (y1 + Y2) * 0.5), 3
                frmMain.view.Circle ((x1 + X2) * 0.5, Y2), 3
                frmMain.view.Circle ((x1 + X2) * 0.5, y1), 3
            End If
        End If
    End If
End Sub
 
 
 
 
 
