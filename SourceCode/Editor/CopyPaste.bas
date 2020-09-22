Attribute VB_Name = "CopyPaste"
Public Sub StoreCurentPosition()
    CopyTo App.Path & "\restore"
End Sub

Public Sub UndoLastMove()
On Error GoTo UndoError
    Close
    CopyTo App.Path & "\restore2"
    DeleteSelected
    PasteFrom App.Path & "\restore", 1
    Kill App.Path & "\restore"
    Name App.Path & "\restore2" As App.Path & "\restore"
    frmMain.OpEdit(1).Caption = "Redo"
    frmMain.OpEdit(1).Enabled = True
    frmMain.DrawModel
UndoError:
End Sub

Public Sub CopyTo(FileName)
    Close
    If Functions.CountSelectedObject = 0 Then
        Exit Sub
    End If
    Dim NowOn As Integer
    Open FileName For Output As #1
    XX = CountSelectedObject
    Print #1, XX
    For NowOn = 1 To cstTotalObjects
        If Object(NowOn).Used = True And Object(NowOn).Selected = True Then
            Print #1, Object(NowOn).EdgeCount
            Print #1, Object(NowOn).FaceCount
            Print #1, Object(NowOn).VertexCount
            Print #1, Object(NowOn).Colour
            For n = 1 To UBound(Object(NowOn).Face())
                Print #1, Object(NowOn).Face(n)
            Next n
            For n = 1 To Object(NowOn).VertexCount
                Print #1, Object(NowOn).Vertex(n).X
                Print #1, Object(NowOn).Vertex(n).Y
                Print #1, Object(NowOn).Vertex(n).Z
                Print #1, Object(NowOn).Vertex(n).TargetName
            Next n
        End If
    Next NowOn
    frmMain.Sbar.Panels(3) = XX & " objects copied..."
'    MsgBox "SAVED"
    Close
End Sub

Public Sub PasteFrom(FileName, Mode)
On Error GoTo NothingToPaste
    Dim NowOn As Integer
    Open FileName For Input As #1
    If Mode = 0 Then
        sz = frmSettings.GridSize
        If ViewMode = 1 Then xo = sz: zo = sz
        If ViewMode = 2 Then xo = sz: yo = sz
        If ViewMode = 3 Then yo = sz: zo = sz
    End If
    Input #1, ObjectsToPaste
    For NowOn = 1 To ObjectsToPaste
        Add = AddObject
        Object(Add).Used = True
        Object(Add).Selected = True
        Input #1, tmp: Object(Add).EdgeCount = tmp
        Input #1, tmp: Object(Add).FaceCount = tmp
        Input #1, tmp: Object(Add).VertexCount = tmp
        Input #1, tmp: Object(Add).Colour = tmp
        ReDim Object(Add).Face(Object(Add).EdgeCount) As Integer
        ReDim Object(Add).Vertex(Object(Add).VertexCount) As VertexDis
        For n = 1 To Object(Add).EdgeCount
            Input #1, edge: Object(Add).Face(n) = edge
        Next n
        For n = 1 To Object(Add).VertexCount
            Input #1, X: Object(Add).Vertex(n).X = X + xo
            Input #1, Y: Object(Add).Vertex(n).Y = Y + yo
            Input #1, Z: Object(Add).Vertex(n).Z = Z + zo
            Input #1, Target: Object(Add).Vertex(n).TargetName = Target
        Next n
        FindOutLine Add
    Next NowOn
    frmMain.Sbar.Panels(3) = ObjectsToPaste & " objects pasted..."
    Close
NothingToPaste:
End Sub


Public Sub PasteFromFile()
On Error GoTo CanceledPasteTo
    frmMain.GetFile.DialogTitle = "Select file to paste from..."
    frmMain.GetFile.FileName = "*.cpy"
    frmMain.GetFile.Filter = "Copy files (*.cpy) |*.cpy|All files (*.*) |*.*"
    frmMain.GetFile.FilterIndex = 1
    
    frmMain.GetFile.ShowOpen
    If frmMain.GetFile.FileName = "" Then Exit Sub
    If Mid(frmMain.GetFile.FileName, 1, 1) = "*" Then Exit Sub
    PasteFrom frmMain.GetFile.FileName, 0
CanceledPasteTo:
End Sub

Public Sub CopyToFile()
On Error GoTo CanceledCopyFrom
    frmMain.GetFile.DialogTitle = "Select file to copy to..."
    frmMain.GetFile.FileName = "*.cpy"
    frmMain.GetFile.Filter = "Copy files (*.cpy) |*.cpy|All files (*.*) |*.*"
    frmMain.GetFile.FilterIndex = 1
    frmMain.GetFile.ShowSave
    If frmMain.GetFile.FileName = "" Then Exit Sub
    If Mid(frmMain.GetFile.FileName, 1, 1) = "*" Then Exit Sub
    lent = Len(frmMain.GetFile.FileName)
    If lent < 5 Then
        frmMain.GetFile.FileName = frmMain.GetFile.FileName + ".cpy"
    Else
        If Mid(frmMain.GetFile.FileName, lent - 3, 1) <> "." Then
            frmMain.GetFile.FileName = frmMain.GetFile.FileName + ".cpy"
        End If
    End If
    CopyTo frmMain.GetFile.FileName
CanceledCopyFrom:
End Sub

