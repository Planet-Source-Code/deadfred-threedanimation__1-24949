Attribute VB_Name = "Functions"
Function FirstObject()
    For N = 1 To cstTotalJoints
        If Object(N).Selected = True Then FirstObject = N: Exit For
    Next N
End Function

Function Destroy(FileName) As Boolean
    Destroy = False
    On Error GoTo NotThisTime_FatBoy
    Kill FileName
    Destroy = True
NotThisTime_FatBoy:
End Function

Public Function SelectFileName(FileType, DialogTitle)
    On Error GoTo CanceledSelect
    frmMain.GetFile.DialogTitle = DialogTitle
    If FileType = "Am5" Then
        frmMain.GetFile.Filter = "AnimationShop files (*.Am5) |*.Am5|AnimationShop Batch files (*.abf) |*.abf|All files (*.*) |*.*"
        frmMain.GetFile.FilterIndex = 1
    End If
    If FileType = "Copy" Then
        frmMain.GetFile.Filter = "Copy files (*.Cpy) |*.cpy|All files (*.*) |*.*"
        frmMain.GetFile.FilterIndex = 1
    End If
    If FileType = "Import" Then
        frmMain.GetFile.Filter = "Compiled data files (*.dat) |*.dat|Terain data files (*.map) |*.map|3D data files (*.asc) |*.asc|Texture images (*.bmp; *.jpg) |*.bmp; *.jpg|All 3D files |*.dat; *.map; *.asc|All files (*.*) |*.*"
        frmMain.GetFile.FilterIndex = 1
    End If
    If FileType = "Compile" Then
        frmMain.GetFile.Filter = "Compiled data files (*.dat) |*.dat|All files (*.*) |*.*"
        frmMain.GetFile.FilterIndex = 1
    End If
    frmMain.GetFile.ShowOpen
    SelectFileName = frmMain.GetFile.FileName
CanceledSelect:
End Function

Public Function SetFileName(FileType, DialogTitle)
    On Error GoTo CanceledSet
    frmMain.GetFile.DialogTitle = DialogTitle
    If FileType = "Copy" Then
        frmMain.GetFile.Filter = "Copy files (*.Cpy) |*.cpy|All files (*.*) |*.*"
        frmMain.GetFile.FilterIndex = 1
    End If
    If FileType = "Am5" Then
        frmMain.GetFile.Filter = "AnimationShop files (*.Am5) |*.Am5|All files (*.*) |*.*"
        frmMain.GetFile.FilterIndex = 1
    End If
    If FileType = "Import" Then
        frmMain.GetFile.Filter = "Compiled data files (*.dat) |*.dat|Terain data files (*.map) |*.map|3D data files (*.asc) |*.asc|Texture images (*.bmp; *.jpg) |*.bmp; *.jpg|All 3D files |*.dat; *.map; *.asc|All files (*.*) |*.*"
        frmMain.GetFile.FilterIndex = 1
    End If
    If FileType = "Compile" Then
        frmMain.GetFile.Filter = "Compiled data files (*.dat) |*.dat|All files (*.*) |*.*"
        frmMain.GetFile.FilterIndex = 1
    End If
    
    
    
    frmMain.GetFile.ShowSave
    SetFileName = frmMain.GetFile.FileName
CanceledSet:
End Function

Public Sub CheckForFrames()
    On Error GoTo notFrames
    Open App.Path & "\frames\56.567" For Output As #1
    Close
    Kill App.Path & "\frames\56.567"
    Exit Sub
notFrames:
    Mes = "To create animation, you need a folder called 'Frames'." & vbNewLine
    Mes = Mes & "This was missing, but has now been created. Please don't remove it."
    MsgBox Mes, vbInformation, "New Folder"
    MkDir App.Path & "\frames\"
    
End Sub

Public Function YouWantToQuit() As Boolean
    If Model.Saved = True Then
        X = MsgBox("Are you sure you want to quit?", vbYesNo + 32)
        YouWantToQuit = False
        If X = 6 Then
            YouWantToQuit = True
            On Error Resume Next
            Destroy App.Path & "\frames\*.dat"
            Destroy App.Path & "\data\Shapes.txt"
            Destroy App.Path & "\restore"
            Destroy App.Path & "\copyfile"
            If SaveSettings = False Then MsgBox "Couldn't save program settings", vbExclamation, "Folder missing"
        End If
    Else
        X = MsgBox("You model has been altered since you last saved" & vbNewLine & "Do you want to save now?", 32 + 3)
        YouWantToQuit = False
        If X = 6 Then
            If Model.ProjectFileName = "*.txt" Then
                FileName = SetFileName("Am5", "Save file...")
                If FileName <> "" Then SaveFile FileName
            Else
                SaveFile Model.ProjectFileName
            End If
            If Model.Saved = False Then Exit Function
            YouWantToQuit = True
            On Error Resume Next
            
            Destroy App.Path & "\frames\*.dat"
            Destroy App.Path & "\data\Shapes.txt"
            Destroy App.Path & "\restore"
            Destroy App.Path & "\copyfile"
            If SaveSettings = False Then MsgBox "Couldn't save program settings", vbExclamation, "Folder missing"
        End If
        If X = 7 Then
            YouWantToQuit = True
            frmMain.tmaLightEffect.Interval = 0
            sceneEdit.Timer1.Interval = 0
        End If
    End If
End Function


Private Function SaveSettings() As Boolean
    On Error GoTo SaveFailed
    SaveSettings = False
    Open App.Path & "\data\Config.dat" For Output As #1
        Print #1, frmSettings.Clours(0).BackColor
        Print #1, frmSettings.Clours(1).BackColor
        Print #1, frmSettings.Clours(2).BackColor
        Print #1, frmSettings.Clours(3).BackColor
        Print #1, frmSettings.Clours(4).BackColor
        Print #1, frmSettings.chkHilighBox
        Print #1, frmSettings.chkNew
        Print #1, frmSettings.DataSpin
        Print #1, frmSettings.GridSize
        Print #1, frmSettings.chkCenter
        If frmMain.OpCustomize(5).Checked = True Then Print #1, 5
        If frmMain.OpCustomize(6).Checked = True Then Print #1, 6
        Print #1, frmSettings.sldGrey
    SaveSettings = True
SaveFailed:
Close
End Function

Function CheckOverwrite(FileName) As Boolean
    CheckOverwrite = False
    On Error GoTo FileEmpty
    Emt = FreeFile
    Open FileName For Input As Emt
    Close Emt
    Resp = MsgBox("This file already exists" & vbNewLine & "Do you want to overwrite it?", vbQuestion + vbYesNo + vbDefaultButton2, "Replace file")
    If Resp = 6 Then GoTo FileEmpty
    Exit Function
FileEmpty:
    CheckOverwrite = True
End Function

Function Snaped(Value)
    If ViewMode > 3 Or frmMain.SnapTo.Checked = False Then
        Snaped = Value
    Else
        Snaped = (CInt(Value / frmSettings.GridSize)) * frmSettings.GridSize
    End If
End Function

Function SetUPSkeliton()
    For N = 1 To cstTotalObjects
        If Object(N).Used = True Then
            For M = 1 To Object(N).VertexCount
                Object(N).Vertex(M).Target = FindTarget(Object(N).Vertex(M).TargetName)
            Next M
        End If
    Next N
    
    ReDim SortJoint(CountJoints) As String
    ReDim RememberW(CountJoints) As Integer
    FF = 0
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True Then
            FF = FF + 1
            RememberW(FF) = N
            SortJoint(FF) = frmMain.Joints.Nodes(BaseFrame(N).Target).FullPath & "\" & BaseFrame(N).Key
        End If
    Next N
    For N = 1 To CountJoints - 1
        For M = 1 To CountJoints - 1
            If Len(SortJoint(M)) > Len(SortJoint(M + 1)) Then
                temp = SortJoint(M)
                SortJoint(M) = SortJoint(M + 1)
                SortJoint(M + 1) = temp
                temp = RememberW(M)
                RememberW(M) = RememberW(M + 1)
                RememberW(M + 1) = temp
            End If
        Next M
    Next N
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True Then
            Num = Num + 1
            KeepRight(Num, 1) = N
            KeepRight(N, 2) = Num
        End If
        XX = XX & KeepRight(N, 1) & " " & KeepRight(N, 2) & vbNewLine
    Next N
End Function

Function JointOver(X, Y) As Integer
    xof = (frmMain.view.ScaleWidth / 2) - frmMain.LRBar
    yof = (frmMain.view.ScaleHeight / 2) - frmMain.UDBar
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True Then
            If ViewMode = 1 Then
                XX = (BaseFrame(N).Position.X * Zoom) + xof
                YY = (BaseFrame(N).Position.Z * Zoom) + yof
            End If
            If ViewMode = 2 Then
                XX = (BaseFrame(N).Position.X * Zoom) + xof
                YY = (BaseFrame(N).Position.Y * Zoom) + yof
            End If
            If ViewMode = 3 Then
                XX = (BaseFrame(N).Position.Z * Zoom) + xof
                YY = (BaseFrame(N).Position.Y * Zoom) + yof
            End If
            If Almostt(X, Y, XX, YY, 5 * Zoom) = True Then
                JointOver = N
            End If
        End If
    Next N
End Function

Function ObjectSelected()
    For N = 1 To cstTotalObjects
        If Object(N).Selected = True Then ObjectSelected = N: Exit Function
    Next N
End Function

Function FlipFaces(NowOn)
    FaceON = 1
    For N = 1 To Object(NowOn).FaceCount
        edge = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
        Startof = FaceON
        ReDim HoldFace(edge)
        For M = 1 To edge
            ThisEdge = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
            HoldFace(edge - (M - 1)) = ThisEdge
        Next M
        FaceON = Startof
        For M = 1 To edge
            Object(NowOn).Face(FaceON) = HoldFace(M): FaceON = FaceON + 1
        Next M
    Next N
End Function


Function FindCenter(Axis As String, NowOnNow As Integer) As Integer
    TotalVetecies = 0
    If NowOnNow = 0 Then
        For NowOn = 1 To cstTotalObjects
            If Object(NowOn).Used = True And Object(NowOn).Selected = True Then
                For N = 1 To Object(NowOn).VertexCount
                    PosX = PosX + Object(NowOn).Vertex(N).X
                    PosY = PosY + Object(NowOn).Vertex(N).Y
                    PosZ = PosZ + Object(NowOn).Vertex(N).Z
                    TotalVetecies = TotalVetecies + 1
                Next N
            End If
        Next NowOn
    Else
        For N = 1 To Object(NowOnNow).VertexCount
            PosX = PosX + Object(NowOnNow).Vertex(N).X
            PosY = PosY + Object(NowOnNow).Vertex(N).Y
            PosZ = PosZ + Object(NowOnNow).Vertex(N).Z
            TotalVetecies = TotalVetecies + 1
        Next N
    End If
    If TotalVetecies = 0 Then Exit Function
    PosX = PosX / TotalVetecies
    PosY = PosY / TotalVetecies
    PosZ = PosZ / TotalVetecies
    If UCase(Axis) = "X" Then FindCenter = PosX
    If UCase(Axis) = "Y" Then FindCenter = PosY
    If UCase(Axis) = "Z" Then FindCenter = PosZ
End Function

Public Function CheckLight(Test)
    CheckLight = True
    For N = 1 To Len(Test)
        X = Asc(UCase(Mid(Test, N, 1)))
        If X < 64 Or X > 91 Then CheckLight = False
    Next N
End Function

Public Function Almostt(x1, y1, x2, y2, Dis) As Boolean
    Almostt = False
    If x1 - Dis < x2 And x1 + Dis > x2 Then
        If y1 - Dis < y2 And y1 + Dis > y2 Then
            Almostt = True
        End If
    End If
End Function

Public Function Edit_Tool()
    For N = 1 To frmMain.TBar2.buttons.Count
        If frmMain.TBar2.buttons(N).Value = tbrPressed Then Edit_Tool = N
    Next N
    If frmMain.opSelectJ = True And frmMain.CurrentSideBar = 3 Then Edit_Tool = 1
    If frmMain.edMode(6) = True And frmMain.CurrentSideBar = 6 Then Edit_Tool = 1
    If frmMain.RotateMove = 1 And frmMain.CurrentSideBar = 4 Then Edit_Tool = 1
    
End Function

Public Function IsThisValid(Number, Mode)
    IsThisValid = False
    On Error GoTo NotNumber
    Number = Number / 2
    If Mode = 0 Then Number = Number / Number
    IsThisValid = True
    Exit Function
NotNumber:
End Function

Public Function GetBrush()
    GetBrush = 1
    For N = 1 To cstTotalObjects
        If Object(N).Used = False Then
            GetBrush = N: Exit Function
        End If
    Next N
End Function

Public Function GetAngle(X, Y)
    If Y = 0 Then
        If X > 0 Then an = 90
        If X < 0 Then an = 270
    Else
        an = Atn(X / Y)
        an = an * Pye
    End If
    If Y < 0 And X > 0 Then an = 180 - (Abs(an))
    If Y < 0 And X < 0 Then an = an + 180
    If Y > 0 And X < 0 Then an = 360 + an
    GetAngle = an
End Function
  
Public Function ShowHelp(HelpTitle)
    ShowHelp = False
    On Error GoTo Bodged
    Open App.Path & "\HelpTopic.txt" For Output As #1
    Print #1, HelpTitle
    Close
    X = App.Path & "\Help.exe"
    Shell X, vbNormalFocus
    ShowHelp = True
    Exit Function
Bodged:
    MsgBox "There was an error staring the help program", , "Error"
End Function

Function AddObject() As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Used = False Then AddObject = N: Exit Function
    Next N
End Function

Function GetJoint() As Integer
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = False Then GetJoint = N: Exit Function
    Next N
End Function

Function FindTarget(Target As String) As Integer
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Key = Target Then
            FindTarget = N: Exit Function
        End If
    Next N
End Function

Function BiggestGroup() As Integer
    Group = 0
    For N = 1 To cstTotalObjects
        If Object(N).Used = True And Object(N).Selected = True Then
            If Object(N).GroupCount > Group Then
                Group = Object(N).Group
            End If
        End If
    Next N
End Function

Function UnusedGroup() As Integer
    Group = 1
    For N = 1 To cstTotalObjects
        If Object(N).Used = True And Object(N).GroupCount <> 0 Then
            If Object(N).Group(Object(N).GroupCount) = Group Then
                Group = Group + 1
            End If
        End If
    Next N
End Function

Function CountObjects() As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Used = True Then
            CountObjects = CountObjects + 1
        End If
    Next N
End Function

Function CountSelectedObject() As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Used = True Then
            If Object(N).Selected = True Then
                Count = Count + 1
            End If
        End If
    Next N
    CountSelectedObject = Count
End Function


Function CountEdges() As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Used = True Then
            CountEdges = CountEdges + Object(N).EdgeCount
        End If
    Next N
End Function

Function CountFaces() As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Used = True Then
            CountFaces = CountFaces + Object(N).FaceCount
        End If
    Next N
End Function

Function CountSelectedFaces() As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Used = True And Object(N).Selected = True Then
            CountSelectedFaces = CountSelectedFaces + Object(N).FaceCount
        End If
    Next N
End Function


Function CountVertecies() As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Used = True Then
            CountVertecies = CountVertecies + Object(N).VertexCount
        End If
    Next N
End Function

Function CountSelectedVertecies() As Integer
    For N = 1 To cstTotalObjects
        If Object(N).Used = True And Object(N).Selected = True Then
            CountSelectedVertecies = CountSelectedVertecies + Object(N).VertexCount
        End If
    Next N
End Function


Function FindScene()
 For N = 1 To cstTotalScenes
    If Scenes(N).Used = False Then FindScene = N: Exit Function
 Next N
End Function

Function CountJoints() As Integer
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True Then CountJoints = CountJoints + 1
    Next N
End Function

Function CountWeapons() As Integer
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True And BaseFrame(N).IsAWeapon = True Then CountWeapons = CountWeapons + 1
    Next N
End Function

        


Function CountSelectedJoints() As Integer
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True And BaseFrame(N).Selected = True Then CountSelectedJoints = CountSelectedJoints + 1
    Next N
End Function


Function CountScenes() As Integer
    For N = 1 To cstTotalScenes
        If Scenes(N).Used = True Then Cs = Cs + 1
    Next N
    CountScenes = Cs
End Function


Function CountFrames() As Integer

End Function

