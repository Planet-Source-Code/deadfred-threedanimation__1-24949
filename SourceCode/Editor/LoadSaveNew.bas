Attribute VB_Name = "LoadSaveNew"

Function LoadFile(FileName) As Boolean
    On Error GoTo ErrorInLoad
    Model.Saved = True
    NewModel 0
    Open FileName For Input As #1
    Input #1, tmp:  Model.ProjectCreated = tmp
    Input #1, tmp:  Model.ProjectFileName = FileName
    Input #1, tmp:  Model.ProjectName = tmp
    Input #1, tmp: Model.StartAnimated = tmp
    Input #1, tmp: Model.StartSceneName = tmp
    Input #1, tmp: Model.StartShowSideBar = tmp
    Input #1, tmp: Model.StartViewMode = tmp
    Input #1, tmp: Model.ShowNotesAtStart = tmp
    Model.ModelNotes = ""
    Input #1, temp
    Do
    Line Input #1, X
    If X <> temp Then Model.ModelNotes = Model.ModelNotes & X & vbNewLine
    Loop While X <> temp
    
    Input #1, Model.BotCost
    Line Input #1, Model.BotDis
    Input #1, Model.BotWeight
    Input #1, Model.BotID
    Input #1, Frco
    If Frco = 0 Then Model.BotForce = False
    If Frco = 1 Then Model.BotForce = True
    
    If Model.StartShowSideBar = True Then frmMain.TBar3.buttons(7).Value = tbrUnpressed
    If Model.StartShowSideBar = False Then frmMain.TBar3.buttons(7).Value = tbrPressed
    frmMain.Caption = App.Title & " [" & Model.ProjectName & "]"
    
    frmMain.MainTab.Tabs(Model.StartViewMode + 1).Selected = True
 
    Input #1, TotalObjectz
    For N = 1 To TotalObjectz
    
        Add = AddObject
        Object(Add).Used = True
        Input #1, tmp: Object(Add).Colour = tmp
        Input #1, tmp: Object(Add).EdgeCount = tmp
        Input #1, tmp: Object(Add).FaceCount = tmp
        Input #1, tmp: Object(Add).GroupCount = tmp
        Input #1, tmp: Object(Add).VertexCount = tmp
        ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
        ReDim Object(Add).Vertex(Object(Add).VertexCount) As VertexDis
        ReDim Object(Add).Group(Object(Add).GroupCount) As Integer
        
        For M = 1 To Object(Add).EdgeCount
            Input #1, Object(Add).Face(M)
        Next M
        For M = 1 To Object(N).VertexCount
                Input #1, tmp: Object(Add).Vertex(M).Target = tmp
                Input #1, tmp: Object(Add).Vertex(M).TargetName = tmp
                Input #1, tmp: Object(Add).Vertex(M).X = tmp
                Input #1, tmp: Object(Add).Vertex(M).Y = tmp
                Input #1, tmp: Object(Add).Vertex(M).Z = tmp
        Next M
        For M = 1 To Object(Add).GroupCount
            Input #1, tmp: Object(Add).Group(M) = tmp
        Next M
        FindOutLine N
 Next N
 Input #1, EndOfObjects
 Input #1, tmp
 Do: Input #1, X
 If X <> "<End of joints>" Then
    N = Val(X)
    Input #1, BaseFrame(N).Colour
    Input #1, BaseFrame(N).Name
    Input #1, BaseFrame(N).Key
    Input #1, BaseFrame(N).Target
    Input #1, BaseFrame(N).Position.X
    Input #1, BaseFrame(N).Position.Y
    Input #1, BaseFrame(N).Position.Z
    
    Input #1, BaseFrame(N).HAngle
    Input #1, BaseFrame(N).Vangle
    Input #1, BaseFrame(N).WeaponName
    Input #1, BaseFrame(N).WeaponType
    
    Input #1, tmp:
    If tmp = 0 Then BaseFrame(N).IsAWeapon = False
    If tmp = 1 Then BaseFrame(N).IsAWeapon = True
    
    Input #1, BaseFrame(N).JointType
    
    Input #1, tmp
    BaseFrame(N).Used = True
    frmMain.Joints.Nodes.Add BaseFrame(N).Target, 4, BaseFrame(N).Key, BaseFrame(N).Name, 4
    frmMain.Joints.Nodes(BaseFrame(N).Key).Tag = 2
    frmMain.Joints.Nodes(BaseFrame(N).Key).EnsureVisible

 End If
 
 Loop While X <> "<End of joints>"
    Input #1, ScenesLoad
    For N = 1 To ScenesLoad
        Input #1, Scenes(N).Name
        Input #1, Scenes(N).Key
        Input #1, Children
        Input #1, Scenes(N).Mode
        Scenes(N).Used = True
        If Scenes(N).Mode = 1 Then
            sceneEdit.Frame.Nodes.Add "Model", 4, Scenes(N).Key, Scenes(N).Name, 2
            sceneEdit.Frame.Nodes(Scenes(N).Key).Tag = 2
        ElseIf Scenes(N).Mode = 2 Then
            sceneEdit.Frame.Nodes.Add "Model", 4, Scenes(N).Key, Scenes(N).Name, 4
            sceneEdit.Frame.Nodes(Scenes(N).Key).Tag = 4
        End If
        sceneEdit.Frame.Nodes(Scenes(N).Key).EnsureVisible
        For M = 1 To Children
            sceneEdit.Frame.Nodes.Add Scenes(N).Key, 4, Scenes(N).Key & "_" & M, "Frame " & M, 1
            sceneEdit.Frame.Nodes(Scenes(N).Key & "_" & M).Tag = 3
        Next M
    Next N

    Do
    Line Input #1, X
    If Mid(X, 1, 1) = "S" Then
        Ti = App.Path & "\frames\" & X
        Open Ti For Output As #2
        Do
            Line Input #1, X
            Print #2, X
        Loop While Mid(X, 1, 1) = "J"
        Close #2
    End If
    Loop While X <> "<End of scenes>"
    Close #1, #2, #3, #4, #5
    frmMain.SortOutScreen
    frmMain.DrawModel
 
    sceneEdit.UpdateToolbarlist
    If Model.ShowNotesAtStart = True And frmMain.Tag <> "Nooo" Then
        AlwaysOnTop frmShowMessage, 1
        frmShowMessage.txtNotes = Model.ModelNotes
        frmShowMessage.Caption = Model.ProjectName
        frmShowMessage.Visible = True
        frmShowMessage.SetFocus
    End If
    DeselectAll
    
    For N = 1 To frmMain.MnuScene.Count - 1
        frmMain.MnuScene(N).Checked = False
    Next N
    If Model.StartSceneName <> 0 Then frmMain.MnuScene(Model.StartSceneName + 1).Checked = True
    
   If Model.StartAnimated = True Then frmMain.StartAnimation
    
    frmMain.CenterView
    frmMain.SBar.Panels(3) = "Model loaded sucessfully"
    LoadFile = True
    Exit Function
    
ErrorInLoad:

Mes = "There was an error while trying to load this file. Make sure that the filename "
Mes = Mes & "and drive are valid, then try again." & vbNewLine & vbNewLine
Mes = Mes & "If the filename is correct, then the file itself may be damaged."
MsgBox Mes, vbCritical, "Error"

End Function

Public Sub SaveFile(FileName)

    On Error GoTo ErrorInSave
    Close
    Open FileName For Output As #1
    Print #1, Model.ProjectCreated
    Print #1, Model.ProjectFileName
    Print #1, Model.ProjectName
    Print #1, Model.StartAnimated
    Print #1, Model.StartSceneName
    Print #1, Model.StartShowSideBar
    Print #1, Model.StartViewMode
    Print #1, Model.ShowNotesAtStart
    Print #1, "--Notes--"
    Print #1, Model.ModelNotes
    Print #1, "--Notes--"
    
    Print #1, Model.BotCost
    Print #1, Model.BotDis
    Print #1, Model.BotWeight
    Print #1, Model.BotID
    If Model.BotForce = False Then Print #1, 0
    If Model.BotForce = True Then Print #1, 1
    
    
    Print #1, CountObjects
    For N = 1 To cstTotalObjects
       If Object(N).Used = True Then
           Print #1, Object(N).Colour
           Print #1, Object(N).EdgeCount
           Print #1, Object(N).FaceCount
           Print #1, Object(N).GroupCount
           Print #1, Object(N).VertexCount
           For M = 1 To UBound(Object(N).Face())
               Print #1, Object(N).Face(M)
           Next M
           For M = 1 To Object(N).VertexCount
                   Print #1, Object(N).Vertex(M).Target
                   Print #1, Object(N).Vertex(M).TargetName
                   Print #1, Object(N).Vertex(M).X
                   Print #1, Object(N).Vertex(M).Y
                   Print #1, Object(N).Vertex(M).Z
           Next M
           For M = 1 To Object(N).GroupCount
               Print #1, Object(N).Group(M)
           Next M
       End If
    Next N
    Print #1, "<End of objects>"
    ReDim SortJoint(CountJoints) As String
    ReDim RememberW(CountJoints) As Integer
    FF = 0
    For N = 1 To cstTotalJoints
        If BaseFrame(N).Used = True Then
            X = BaseFrame(N).Target
            XX = frmMain.Joints.Nodes(X).FullPath & "\" & BaseFrame(N).Key
            FF = FF + 1
            RememberW(FF) = N
            SortJoint(FF) = XX
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

    Print #1, "-----------"
    For N = 1 To CountJoints
        M = RememberW(N)
        Print #1, M
        Print #1, BaseFrame(M).Colour
        Print #1, BaseFrame(M).Name
        Print #1, BaseFrame(M).Key
        Print #1, BaseFrame(M).Target
        Print #1, BaseFrame(M).Position.X
        Print #1, BaseFrame(M).Position.Y
        Print #1, BaseFrame(M).Position.Z
        Print #1, BaseFrame(M).HAngle
        Print #1, BaseFrame(M).Vangle
        Print #1, BaseFrame(M).WeaponName
        Print #1, BaseFrame(M).WeaponType
        If BaseFrame(M).IsAWeapon = True Then Print #1, 1
        If BaseFrame(M).IsAWeapon = False Then Print #1, 0
        Print #1, BaseFrame(M).JointType
        Print #1, "-----------"
    Next N
    
    Print #1, "<End of joints>"
    Print #1, CountScenes
    For N = 1 To cstTotalScenes
        If Scenes(N).Used = True Then
            Print #1, Scenes(N).Name; ",  ";
            Print #1, Scenes(N).Key; ",  ";
            Print #1, sceneEdit.Frame.Nodes(Scenes(N).Key).Children
            Print #1, Scenes(N).Mode
        End If
    Next N
    Print #1, "<End of scene info>"
    
    For N = 1 To cstTotalScenes
        If Scenes(N).Used = True Then
            For M = 1 To sceneEdit.Frame.Nodes(Scenes(N).Key).Children
                GetFile = Scenes(N).Key & "_" & M & ".dat"
                Print #1, GetFile
                GetFile = App.Path & "\frames\" & GetFile
                Open GetFile For Input As #2
                    Do:  Line Input #2, hole:  Print #1, hole
                    Loop While EOF(2) <> True
                Close #2
            Next M
        End If
    Next N
    Print #1, "<End of scenes>"
    Close #1, #2, #3, #4, #5
    Model.ProjectFileName = FileName
    Model.Saved = True
    frmMain.SBar.Panels(3) = "Model saved " & Time
    frmMain.Caption = App.Title & " [" & Model.ProjectName & "]"
    Exit Sub
ErrorInSave:

Mes = "There was an error while trying to save this file. Make sure that the filename "
Mes = Mes & "and drive are valid, then try again." & vbNewLine & vbNewLine
Mes = Mes & "If you quit now, your changes will not be saved."
MsgBox Mes, vbCritical, "Error"

End Sub

Public Sub NewModel(Mode)
 frmMain.ShapeList.Selected(0) = True
 If Mode = 1 Then
  X = MsgBox("Are you sure you want to start a new model?", vbYesNo + 32)
  If X = 7 Then Exit Sub
 End If

 Model.ProjectFileName = "*.txt"
 Model.ProjectName = "Untitled"
  
 If frmSettings.chkNew = 1 Then
    If Mode = 1 Then Model.ProjectName = InputBox("Enter a name for this model", "Enter name", "Untitled")
    If Model.ProjectName = "" Then Model.ProjectName = "Unititled"
 End If
 
 Model.StartShowSideBar = True
 Model.BotDis = ""
 Model.BotCost = 0
 Model.BotWeight = 0
 Model.ModelNotes = ""
 Model.StartAnimated = False
 Model.StartViewMode = 0
 Model.StartSceneName = 0
 Model.ProjectCreated = Date$
 Model.Saved = False
 
 For N = 1 To cstTotalObjects: Object(N).Used = False
 Object(N).Selected = False
 Object(N).Colour = 0
 Next N
 For N = 1 To cstTotalJoints
    BaseFrame(N).Used = False
    BaseFrame(N).Position.X = 0
    BaseFrame(N).Position.Y = 0
    BaseFrame(N).Position.Z = 0
    BaseFrame(N).Name = ""
    BaseFrame(N).Target = ""
    
 Next N
 For N = 1 To cstTotalScenes: Scenes(N).Used = False: Next N
 frmMain.Caption = App.Title & " [" & Model.ProjectName & "]"
 frmMain.Joints.Nodes.Clear
 sceneEdit.Frame.Nodes.Clear
 frmMain.Joints.Nodes.Add , , "Model", Model.ProjectName, 5
 frmMain.Joints.Nodes("Model").Tag = 1
 sceneEdit.Frame.Nodes.Add , 1, "Model", "Model", 3
 sceneEdit.Frame.Nodes("Model").Tag = 1
 frmMain.Joints.Nodes("Model").Selected = True
 frmMain.Caption = App.Title & " [" & Model.ProjectName & "]"
 frmMain.CurrentSideBar = 5
 For N = 1 To cstTotalObjects
    Object(N).Used = False
 Next N
 frmMain.TBar2.buttons(1).Value = tbrPressed
 frmMain.DrawModel
 Model.Saved = True
End Sub

Public Sub EditObjectProperties()
    If CountSelectedObject = 0 And CountSelectedJoints <> 0 Then
        frmJointProp.Visible = True
        frmMain.Enabled = False
        frmJointProp.SetFocus
    End If
    
    If CountSelectedObject <> 0 And CountSelectedJoints = 0 Then
        frmObjectPrp.Visible = True
        frmMain.Enabled = False
        frmObjectPrp.SetFocus
        frmObjectPrp.RunAtStart
    End If
    
    If CountSelectedObject <> 0 And CountSelectedJoints <> 0 Then
        frmMain.SBar.Panels(3) = "Make sure that joints and objects are not selected together..."
    End If
    
End Sub

Public Sub ExportModel()
    If CountObjects = 0 Then
        MsgBox "You cannot export an empty model", 48
        Exit Sub
    End If
    frmMain.Enabled = False
    FrmExport.Visible = True
End Sub

Public Sub ModelProperties()
 frmMain.Enabled = False
 frmProperties.Visible = True
End Sub








