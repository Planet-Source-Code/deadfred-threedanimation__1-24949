VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_Nodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CLS_Node"
Attribute VB_Ext_KEY = "Member0" ,"CLS_Node"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variable locale pour la collection
Private mCol As Collection

Public Function Add(ByVal Text As String, _
            ByVal Index As Long, _
            ByVal Tag As String, _
            ByVal Parent As String, _
            ByVal Key As String, _
            ByVal Image As Long, _
            ByVal Selectedimage As Long, _
            Optional ByVal sKey As String) As CLS_Node
        
    'créer un nouvel objet
    Dim objNewMember As CLS_Node
    Set objNewMember = New CLS_Node


    'définir les propriétés passées à la méthode
    objNewMember.Selectedimage = Selectedimage
    objNewMember.Image = Image
    objNewMember.Tag = Tag
    objNewMember.Text = Text
    objNewMember.Index = Index
    objNewMember.Parent = Parent
    objNewMember.Key = Key
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'renvoyer l'objet créé
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CLS_Node
Attribute Item.VB_UserMemId = 0
    'Utilisé pour référencer un élément dans la collection
    'vntIndexKey contient l'index ou la clé dans la collection
    'Il est donc déclaré en tant que Variant
    'Syntaxe: Set toto = x.Item(xyz) ou Set toto = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'utilisée pour lire le nombre d'éléments dans la
    'collection. Syntaxe: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'utilisée lors de la suppression d'un élément de la collection
    'vntIndexKey contient l'index ou la clé, il est donc
    'déclaré en tant que Variant
    'Syntaxe: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Cette propriété vous permet d'énumérer
    'cette collection avec la syntaxe For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Crée la collection lorsque cette classe est créée
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'Détruit la collection lorsque cette classe est détruite
    Set mCol = Nothing
End Sub

