Attribute VB_Name = "Compile"
Public Sub SingleFaceExport()
    If FrmExport.chkNotes = 1 Then notes = True
    If notes = True Then Print #1, "// This is the number of seperate faces in the file..."
    Print #1, CountFaces
    If notes = True Then
        Print #1, "// The first number in each line is the number of corners that make up the face"
        Print #1, "// The other numbers on the line come in groups of three, and locate a 3D point in space."
        Print #1, "// By joining up these points in order, you outline the face"
    End If
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True Then
            FaceON = 1
            For n = 1 To Object(NowOn).FaceCount
                EdgeCount = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
                Print #1, EdgeCount;
                For m = 1 To EdgeCount
                    EdgeOn = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
                    Print #1, "  , "; Object(NowOn).Vertex(EdgeOn).X;
                    Print #1, ", "; Object(NowOn).Vertex(EdgeOn).Y;
                    Print #1, ", "; Object(NowOn).Vertex(EdgeOn).Z;
                Next m
                Print #1,
            Next n
        End If
    Next NowOn
End Sub

Public Sub SingleObjectMultiLineExport()
    FirstNotes = False
    If FrmExport.chkNotes = 1 Then notes = True
    If FrmExport.chkStartAt.Value = 0 Then MoveBack = 1
    If notes = True Then
        Print #1, "// This number is how many objects there are in the model. Each object is a 3D shape, such as"
        Print #1, "// cube, dimond or ball etc."
    End If
    Print #1, CountObjects
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True Then
            If notes = True And FirstNotes = False Then
                Print #1, "// The next two numbers record the number of vertecies (3D points in space) in the first"
                Print #1, "// object, and then number of edges (Lines that join up the vertecies to define a face)"
                Print #1, "// in the object. The third number is the colour of the object, in hexedecimal"
            End If
            Print #1, Object(NowOn).VertexCount
            Print #1, Object(NowOn).EdgeCount
            Print #1, Object(NowOn).Colour
            If notes = True And FirstNotes = False Then
                Print #1, "// This is a list of all the vertecies. You will need to load these into an array, and then"
                Print #1, "// rotate them around to spin the object"
            End If
            For n = 1 To Object(NowOn).VertexCount
                Print #1, Object(NowOn).Vertex(n).X; ", ";
                Print #1, Object(NowOn).Vertex(n).Y; ", ";
                Print #1, Object(NowOn).Vertex(n).Z
            Next n
            If notes = True And FirstNotes = False Then
                Print #1, "// This is a list of the faces. The first number in the list states how many corners each face"
                Print #1, "//  has. The rest of the line refers to vertecies in the list above. To draw a face, rotate the"
                Print #1, "//  list of vertecies above, and then go throught each of the lines below, and draw a line"
                Print #1, "//  from the position of vertex '1' to vertex '2' and so on until you get to the end of the"
                Print #1, "//  line, and then draw a line back to the first vertex in the line..."
            End If
            FaceON = 1
            For n = 1 To Object(NowOn).FaceCount
                EdgeCount = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
                Print #1, EdgeCount; ", ";
                For m = 1 To EdgeCount
                    EdgeOn = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
                    Print #1, EdgeOn - MoveBack;
                    If m <> EdgeCount Then Print #1, ", ";
                Next m
                Print #1,
            Next n
            If notes = True And FirstNotes = False Then
                Print #1, "// That is all the info for the first object. Any data below is for more objects, and is stored"
                Print #1, "// in the same format as is described above..'"
            End If
            Print #1,
            FirstNotes = True
        End If
    Next NowOn
End Sub

Public Sub SingleObjectExport()
    If FrmExport.chkNotes = 1 Then notes = True
    If FrmExport.chkStartAt.Value = 0 Then MoveBack = 1
    If notes = True Then
        Print #1, "// This number is how many objects there are in the model. Each object is a 3D shape, such as"
        Print #1, "// cube, dimond or ball etc."
    End If
    Print #1, CountObjects
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True Then
            If notes = True And FirstNotes = False Then
                Print #1, "// The next two numbers record the number of vertecies (3D points in space) in the first"
                Print #1, "// object, and then number of edges (Lines that join up the vertecies to define a face)"
                Print #1, "// in the object. The third number is the colour of the object, in hexedecimal"
            End If
            Print #1, Object(NowOn).VertexCount
            Print #1, Object(NowOn).EdgeCount
            Print #1, Object(NowOn).Colour
            If notes = True And FirstNotes = False Then
                Print #1, "// This is a list of all the vertecies. You will need to load these into an array, and then"
                Print #1, "// rotate them around to spin the object"
            End If
            For n = 1 To Object(NowOn).VertexCount
                Print #1, Object(NowOn).Vertex(n).X; ", ";
                Print #1, Object(NowOn).Vertex(n).Y; ", ";
                Print #1, Object(NowOn).Vertex(n).Z
            Next n
            If notes = True And FirstNotes = False Then
                Print #1, "// This is a list of the faces. The first number in the list states how many corners the"
                Print #1, "// first face has. When you have read that many numbers into the list, the next number"
                Print #1, "// you will read will be the number of corners on the next face. All info for all the faces"
                Print #1, "// are stored in this one list."
            End If
            For n = 1 To Object(NowOn).EdgeCount
                Print #1, Object(NowOn).Face(n) - MoveBack;
                If n <> Object(NowOn).EdgeCount Then Print #1, ", ";
            Next n
            Print #1,
            If notes = True And FirstNotes = False Then
                Print #1, "// That is all the info for the first object. Any data below is for more objects, and is stored"
                Print #1, "// in the same format as is described above..'"
            End If
            Print #1,
            FirstNotes = True
        End If
    Next NowOn
End Sub

Public Sub CompleteObject()
    If FrmExport.chkStartAt.Value = 0 Then MoveBack = 1
    If FrmExport.chkNotes = 1 Then notes = True
    FrmExport.List1.Clear
    Timei.TimeLeft.Max = CountFaces
    Timei.TimeLeft = 0
    Timei.Refresh
    For n = 1 To cstTotalObjects
        If Object(n).Used = True Then
            For m = 1 To Object(n).VertexCount
                If FrmExport.Mode(4) = True Or FrmExport.Mode(6) = True Then
                    X = Object(n).Vertex(m).X & ", " & Object(n).Vertex(m).Y & ", " & Object(n).Vertex(m).Z & ", " & Object(n).Vertex(m).Target
                Else
                    X = Object(n).Vertex(m).X & ", " & Object(n).Vertex(m).Y & ", " & Object(n).Vertex(m).Z
                End If
                FrmExport.List1.AddItem X
            Next m
        End If
    Next n
    rCount = 0
Restart:
    For n = 0 To FrmExport.List1.ListCount - 2
        If FrmExport.List1.List(n) = FrmExport.List1.List(n + 1) Then
            FrmExport.List1.RemoveItem n + 1
            rCount = rCount + 1
            GoTo Restart
        End If
    Next n
    DoEvents
'    MsgBox rCount & " items removed..."
    If notes = True Then
        Print #1, "// This is the number of vertecies in the model. Repeated vertecies have been"
        Print #1, "// removed, so the model is as efficient as possible."
        Print #1, "// The next number is the number of faces in the model."
    End If
    Print #1, FrmExport.List1.ListCount
    If FrmExport.Mode(1) = True Or FrmExport.Mode(4) = True Then
        Print #1, CountEdges
    Else
        Print #1, CountFaces
    End If
    If notes = True Then
        Print #1, "// This list contains all the different vertecies. Each vertex is held as three"
        Print #1, "// numbers, indicating the X, Y and Z axies."
        If FrmExport.Mode(4) = True Or FrmExport.Mode(6) = True Then
            Print #1, "// The fourth number is the number of the joint that that vertex is attached to."
        End If
    End If
    For n = 0 To FrmExport.List1.ListCount - 1
        Print #1, FrmExport.List1.List(n)
    Next n
    If notes = True Then
        If FrmExport.Mode(1) = True Or FrmExport.Mode(4) = True Then
            Print #1, "// This is a list of the faces. The first number in the list states how many corners the"
            Print #1, "// first face has. When you have read that many numbers into the list, the next number"
            Print #1, "// you will read will be the number of corners on the next face. All info for all the faces"
            Print #1, "// are stored in this one list."
        Else
            Print #1, "// This is the list holding each face. Each line represents one face."
            Print #1, "// The first number is the number of edges in the face. The rest of the line is a list"
            Print #1, "// of vertecies that, when joined together, make the outline of the face."
        End If
    End If
        
        
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True Then
            FaceON = 1
            For n = 1 To Object(NowOn).FaceCount
                EdgeCount = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
                If FirstCommaTime = True Then
                    If FrmExport.Mode(1) = False And FrmExport.Mode(4) = False Then Print #1, ", ";
                    FirstCommaTime = True
                End If
                
                Print #1, EdgeCount; ", ";
                For m = 1 To EdgeCount
                    EdgeOn = Object(NowOn).Face(FaceON): FaceON = FaceON + 1
                    If FrmExport.Mode(4) = True Or FrmExport.Mode(6) = True Then
                        X = Object(NowOn).Vertex(EdgeOn).X & ", " & Object(NowOn).Vertex(EdgeOn).Y & ", " & Object(NowOn).Vertex(EdgeOn).Z & ", " & Object(NowOn).Vertex(EdgeOn).Target
                    Else
                        X = Object(NowOn).Vertex(EdgeOn).X & ", " & Object(NowOn).Vertex(EdgeOn).Y & ", " & Object(NowOn).Vertex(EdgeOn).Z
                    End If
                    For FF = 0 To FrmExport.List1.ListCount
                        If FrmExport.List1.List(FF) = X Then
                            UniqueNumber = FF: Exit For
                        End If
                    Next FF
                    Print #1, UniqueNumber - MoveBack + 1;
                    If m <> EdgeCount Then Print #1, ", ";
                Next m
                If FrmExport.Mode(1) = True Or FrmExport.Mode(4) = True Then
                Else
                   Print #1,
                End If
                Timei.TimeLeft = Timei.TimeLeft + 1
            Next n
        End If
    Next NowOn
    If FrmExport.Mode(4) = True Or FrmExport.Mode(6) = True Then
        If FrmExport.Mode(1) = True Or FrmExport.Mode(4) = True Then
            Print #1,
        End If
        If notes = True Then
            Print #1, "// This is the list of joints in the model. Each row is one joint. The first thee numbers"
            Print #1, "// hold the location of the joint. The next number is the number of the joint. The number "
            Print #1, "// after is the number of the joints target. The string is the name of the joint."
            Print #1, "// The joints are in order, so that a joint always comes after its target in the list."
            Print #1, "// The first number is the number of joints. The second number states whether the"
            Print #1, "// skeliton should be drawn as part of the model, then the actual list starts."
        End If
        If Model.BotForce = False Then Print #1, 0
        If Model.BotForce = True Then Print #1, 1
        Print #1, CountJoints
        For n = 1 To CountJoints
            NowSave = RememberW(n)
            If BaseFrame(NowSave).JointType = "" Then BaseFrame(NowSave).JointType = "None"
            Print #1, BaseFrame(NowSave).Position.X; ", ";
            Print #1, BaseFrame(NowSave).Position.Y; ", ";
            Print #1, BaseFrame(NowSave).Position.Z; ", ";
            Print #1, NowSave; ", ";
            Print #1, FindTarget(BaseFrame(NowSave).Target); ", "; Chr$(34);
            Print #1, BaseFrame(NowSave).Name; Chr$(34); ", "; Chr$(34);
            Print #1, BaseFrame(NowSave).JointType; Chr$(34)
            
        Next n
    End If
    If FrmExport.chkWeapon = 1 Then SaveWeapons notes
End Sub


Private Sub SaveWeapons(notes)
    If notes = True Then
        Print #1, "// This is the list of weapons in the model. Each row is one weapon. The first number"
        Print #1, "// is the number of the joint where this weapon is located. The first string is the name"
        Print #1, "// of the weapon. The second string is the type of weapon that the weapon is. The next"
        Print #1, "// two numbers are the horizontal and vertical angles that te weaopn fires in"
        Print #1, "// The first number is the number of weapons. Then the actual list starts."
    End If
    Print #1, CountWeapons
    For n = 1 To cstTotalJoints
        If BaseFrame(n).Used = True And BaseFrame(n).IsAWeapon = True Then
            Print #1, n; ", "; Chr$(34);
            Print #1, BaseFrame(n).WeaponName; Chr$(34); " , "; Chr$(34);
            Print #1, BaseFrame(n).WeaponType; Chr$(34); " ,";
            Print #1, BaseFrame(n).HAngle; ",";
            Print #1, BaseFrame(n).VAngle
        End If
    Next n
End Sub






