Attribute VB_Name = "ImportStuff"
Sub LoadBatchFile(FileName)
    Dim LineON As Integer
    Open FileName For Input As #10
        Do While EOF(10) = False
            Line Input #10, instruct
            If Len(instruct) > 2 Then
                If Left(instruct, 2) <> "//" Then
                    comm = Trim(Left(instruct, InStr(1, instruct, ":") - 1))
                    opper = Trim(Mid(instruct, InStr(1, instruct, ":") + 1, 99999))
                    LineON = LineON + 1
                    Select Case UCase(comm)
                        Case "ECHO"
                            MsgBox opper
                            
                        Case "LOAD"
                            If Mid(opper, 2, 2) = ":\" Then
                                If LoadFile(opper) = False Then
                                    MsgBox "File " & opper & " failed to load"
                                    frmMain.Sbar.Panels(3) = "####### Error - Couldn't complete bath file (" & LineON & ") #######"
                                    Exit Sub
                                End If
                            Else
                                If LoadFile(App.Path & "\" & opper) = False Then
                                    MsgBox "File " & App.Path & "\" & opper & " failed to load"
                                    frmMain.Sbar.Panels(3) = "####### Error - Couldn't complete bath file (" & LineON & ") #######"
                                    Exit Sub
                                End If
                            End If
                            
                        Case "SETVIEW"
                            XX = Val(opper)
                            If XX > 0 And XX < 7 Then
                                frmMain.MainTab.Tabs(XX).Selected = True
                            End If
                            
                        Case "NEW"
                            NewModel 0
                            
                        Case "OPENCOMPILER"
                            FrmExport.Visible = True
                            frmMain.Enabled = False
                            
                        Case "COMPILEMODE"
                            XX = Val(opper)
                            If XX = 1 Then FrmExport.Mode(0) = True
                            If XX = 2 Then FrmExport.Mode(2) = True
                            If XX = 3 Then FrmExport.Mode(3) = True
                            If XX = 4 Then FrmExport.Mode(1) = True
                            If XX = 5 Then FrmExport.Mode(5) = True
                            If XX = 6 Then FrmExport.Mode(4) = True
                            If XX = 7 Then FrmExport.Mode(6) = True
                            
                        Case "COMPILE"
                            If opper = "" Then
                                XX = Mid(Model.ProjectFileName, 1, Len(Model.ProjectFileName) - 4) & ".dat"
                                FrmExport.CompilerControler XX, 0
                            ElseIf Mid(opper, 2, 2) = ":\" Then
                                FrmExport.CompilerControler opper, 0
                            Else
                                FrmExport.CompilerControler App.Path & "\" & opper, 0
                            End If
                        
                        
                    End Select
                End If
            End If
        Loop
    Close
End Sub


Sub ImportModel(FileName)
    If Left(FileName, 1) = "*" Then Exit Sub
    FileType = Right(FileName, 3)
    If Mid(FileName, Len(FileName) - 3, 1) <> "." Then
        Ms = "This file has no extension, and cannot be identified. "
        Ms = Ms & "Please give it an extension occording to what format the file in in. "
        Ms = Ms & "Avalible formats are " & ImportTypes & vbNewLine & vbNewLine
        Ms = Ms & App.Title & " will try to open this file now as thought it had a TXT extenstion, but it may not work if this is the wrong format" & vbNewLine
        MsgBox Ms, 48
        FileType = "txt"
    End If
    FileType = LCase(FileType)
    If FileType <> "bmp" Then
        Open FileName For Input As #1
    End If
    Add = AddObject
    frmMain.Sbar.Panels(3) = "####### Error - Couldn't import data file #######"
    
    Select Case FileType
        Case "bmp", "jpg"
            frmMain.Texture.Picture = LoadPicture(FileName)
            frmMain.Sbar.Panels(3) = "Image loaded into the texture window"
        Case "map"
            Input #1, XX
            Input #1, YY
            FaceCount = ((XX - 1) * (YY - 1)) * 2
            ReDim Object(Add).Vertex(XX * YY) As VertexDis
            ReDim Object(Add).Face(FaceCount * 4) As Integer
            VertexOn = 1
            For n = 1 To XX
                For m = 1 To YY
                    Input #1, hyte
                    Object(Add).Vertex(VertexOn).X = (n * 100) - (XX * 50) - 50
                    Object(Add).Vertex(VertexOn).Y = -hyte
                    Object(Add).Vertex(VertexOn).Z = (m * 100) - (YY * 50) - 50
                    VertexOn = VertexOn + 1
                Next m
            Next n
            FaceON = 1
            For n = 1 To XX - 1
                For m = 1 To YY - 1
                    xx3 = m + ((n - 1) * XX)
                    xx2 = m + ((n - 1) * XX) + 1
                    xx1 = m + ((n - 1) * XX) + XX
                    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
                    Object(Add).Face(FaceON) = xx1: FaceON = FaceON + 1
                    Object(Add).Face(FaceON) = xx2: FaceON = FaceON + 1
                    Object(Add).Face(FaceON) = xx3: FaceON = FaceON + 1
                    xx3 = m + ((n - 1) * XX) + XX + 1
                    xx2 = m + ((n - 1) * XX) + 1
                    xx1 = m + ((n - 1) * XX) + XX
                    Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
                    Object(Add).Face(FaceON) = xx3: FaceON = FaceON + 1
                    Object(Add).Face(FaceON) = xx2: FaceON = FaceON + 1
                    Object(Add).Face(FaceON) = xx1: FaceON = FaceON + 1
                Next m
            Next n
            Object(Add).EdgeCount = FaceCount * 4
            Object(Add).FaceCount = FaceCount
            Object(Add).VertexCount = XX * YY
            frmMain.Sbar.Panels(3) = "MAP file imported - " & XX * YY & " vertecies - " & FaceCount & " faces"
            Object(Add).Used = True
            FindOutLine (Add)
        Case "dat"
            Input #1, temp, Model.BotID
            Input #1, temp, Model.BotDis
            Input #1, temp, Model.BotCost
            Input #1, temp, Model.BotWeight
            Input #1, VertexCount
            Input #1, FaceCount

            ReDim Object(Add).Vertex(VertexCount) As VertexDis
            Object(Add).VertexCount = VertexCount
            For n = 1 To VertexCount
                Input #1, X
                Input #1, Y
                Input #1, Z
                Object(Add).Vertex(n).X = -X
                Object(Add).Vertex(n).Y = Z
                Object(Add).Vertex(n).Z = Y
                
                Input #1, Targ '##############################
                Object(Add).Vertex(n).Target = Targ '###############
                
            Next n
            
            Object(Add).FaceCount = FaceCount
            Object(Add).EdgeCount = EdgeCount
            
      
            
            FaceON = 1
            For n = 1 To FaceCount
                
                Input #1, Edges
                ReDim Preserve Object(Add).Face(FaceON)
                Object(Add).Face(FaceON) = Edges
                FaceON = FaceON + 1
                
                For m = 1 To Edges
                    
                    Input #1, FaceCorner
                    ReDim Preserve Object(Add).Face(FaceON)
                    Object(Add).Face(FaceON) = FaceCorner + 1
                    FaceON = FaceON + 1
                    
                 Next m
            Next n
            
            
            Object(Add).EdgeCount = FaceON
            
            
            frmMain.Sbar.Panels(3) = "DAT file imported - " & VertexCount & " vertecies - " & FaceCount & " faces"
            Object(Add).Used = True
            FindOutLine (Add)
         Case "asc"
            ReDim poss(12)
            Do
            Input #1, lne
            Loop While lne <> "Tri-mesh"
            Input #1, LineON
            Num = 1
            For n = 1 To Len(LineON)
                If Mid(LineON, n, 1) = ":" Then
                If Num = 1 Then VertexCount = Mid(LineON, n + 1, 4)
                If Num = 2 Then FaceCount = Mid(LineON, n + 1, 4)
                Num = Num + 1
                End If
            Next n
            Input #1, LineON
            ReDim Object(Add).Vertex(VertexCount) As VertexDis
            ReDim Object(Add).Face(FaceCount * 4) As Integer
            Object(Add).FaceCount = FaceCount
            Object(Add).EdgeCount = FaceCount * 4
            Object(Add).VertexCount = VertexCount
            For m = 1 To VertexCount
                Input #1, LineON
                Num = 1
                For n = 1 To Len(LineON)
                    If Mid(LineON, n, 1) = ":" Then
                    poss(Num) = n
                    Num = Num + 1
                    End If
                Next n
                poss(5) = Len(LineON)
                X = Val(Mid(LineON, poss(2) + 1, poss(3) - poss(2) - 2))
                Y = Val(Mid(LineON, poss(3) + 1, poss(4) - poss(3) - 2))
                Z = Val(Mid(LineON, poss(4) + 1, poss(5) - poss(4) - 2))
                Object(Add).Vertex(m).X = X * 2
                Object(Add).Vertex(m).Y = -Z * 2
                Object(Add).Vertex(m).Z = Y * 2
            Next m
            Input #1, LineON
            FaceON = 1
            For m = 1 To FaceCount
                Input #1, LineON
                Num = 1
                For n = 1 To Len(LineON)
                    If Mid(LineON, n, 1) = ":" Then
                    poss(Num) = n
                    Num = Num + 1
                    End If
                Next n
                X = Val(Mid(LineON, poss(2) + 1, poss(3) - poss(2) - 2)) + 1
                Y = Val(Mid(LineON, poss(3) + 1, poss(4) - poss(3) - 2)) + 1
                Z = Val(Mid(LineON, poss(4) + 1, poss(5) - poss(4) - 3)) + 1
                Object(Add).Face(FaceON) = 3: FaceON = FaceON + 1
                Object(Add).Face(FaceON) = X: FaceON = FaceON + 1
                Object(Add).Face(FaceON) = Y: FaceON = FaceON + 1
                Object(Add).Face(FaceON) = Z: FaceON = FaceON + 1
            Next m
            Object(Add).Used = True
            FindOutLine (Add)
            frmMain.Sbar.Panels(3) = "ASC file imported - " & VertexCount & " vertecies - " & FaceCount & " faces"
        Case "dfx"
        Case "pov"
        Case Else
            Ms = "This file has an unknown extension, and cannot be identified" & vbNewLine
            Ms = Ms & "Avalible formats are - .BMP .JPG .ASC .POV .DFX .DAT .MAP" & vbNewLine
            MsgBox Ms, 16
    End Select
CancelInput:
    frmMain.View.AutoRedraw = True
    frmMain.View.DrawStyle = 0
    frmMain.DrawModel
    Close #1, #2, #3, #4, #5
End Sub
