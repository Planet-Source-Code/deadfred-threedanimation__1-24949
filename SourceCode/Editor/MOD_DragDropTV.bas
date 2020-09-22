Attribute VB_Name = "MOD_DragDropTV"
Option Explicit




Private ChildNodes As CLS_Nodes

'==================================================================================
'   Make a Drag & Drop with NODES in MEMORY
'==================================================================================
Public Function DragAndDrop(TV As TreeView, _
            NodeDrag As Node, _
            NodeDrop As Node) As Boolean

    On Error GoTo gestion_erreur

    Dim i As Long
    Dim Node As CLS_Node
    
    DragAndDrop = True

    Set ChildNodes = New CLS_Nodes

    If NodeDrag.Index <> NodeDrop.Index Then
    
        If FindNodeBelow(NodeDrag, NodeDrop, ChildNodes) Then

            '   CLEAR NodeDrag
            TV.Nodes.Remove NodeDrag.Index

            '   MOVE  NodeDrag
            TV.Nodes.Add NodeDrop, tvwChild, ChildNodes(1).Key, ChildNodes(1).Text, ChildNodes(1).Image, ChildNodes(1).Selectedimage
            TV.Nodes(ChildNodes(1).Key).Tag = 2
            ChildNodes.Remove 1
            
            For Each Node In ChildNodes
                With Node
                    TV.Nodes.Add .Parent, tvwChild, .Key, .Text, .Image, .Selectedimage
                End With
                TV.Nodes(Node.Key).Tag = 2
            Next
            
        Else
            ' CAN NOT MOVE NODE IF NodeDrop is in ChildNodes of NodeDrag
            DragAndDrop = False
        End If
    Else
        DragAndDrop = False
    End If

    Set ChildNodes = Nothing
    Exit Function

gestion_erreur:
    DragAndDrop = False
    MsgBox Err.Description, vbCritical, "ERREUR n°" & CStr(Err.Number)
End Function

'==================================================================================
'   Save all Children NODE of StartNode in a Collection
'==================================================================================
Public Function FindNodeBelow(StartNode As Node, _
        ForbidenNode As Node, _
        branche As CLS_Nodes) As Boolean

    Dim Node As Node
    Dim Continue As Boolean

    On Error GoTo gestion_erreur

    FindNodeBelow = True
    
    With StartNode
        branche.Add .Text, .Index, .Tag, "", .Key, .Image, .Selectedimage, .Key
    End With
    
    Set Node = StartNode
    Continue = True
    
    ' si le node de départ a des fils
    If StartNode.Children <> 0 Then
    
        While Continue
        
            ' si le node actuel a des fils on descend d'un niveau
            If Node.Children <> 0 Then
                Set Node = Node.Child
            Else
                ' si c'est le dernier node attaché on remonte
                If Node = Node.LastSibling Then
                
                    ' on remonte tant que necessaire
                    While Node = Node.LastSibling And Continue
                    
                        ' test pour ne pas trop remonté
                        If Node.Parent <> StartNode Then
                            Set Node = Node.Parent
                        Else
                            Continue = False
                        End If
                        
                    Wend
                    
                End If
                ' on passe au node suivant
                If Continue Then Set Node = Node.Next
                
            End If
            
            ' il est impossible de deplacer un node vers l'une de ces branches
            If Node = ForbidenNode Then
                ' Drag & Drop impossible
                FindNodeBelow = False
                Continue = False
            End If
            
            If Continue Then
                ' on ajoute l'image du node a la collection
                With Node
                    branche.Add .Text, .Index, .Tag, .Parent.Key, .Key, .Image, .Selectedimage, .Key
                End With
            End If
            
        Wend

    End If
    
    Exit Function

gestion_erreur:
    MsgBox Err.Description, vbCritical, "ERREUR n°" & CStr(Err.Number)

End Function


