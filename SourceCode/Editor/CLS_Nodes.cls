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
        
    'cr�er un nouvel objet
    Dim objNewMember As CLS_Node
    Set objNewMember = New CLS_Node


    'd�finir les propri�t�s pass�es � la m�thode
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


    'renvoyer l'objet cr��
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CLS_Node
Attribute Item.VB_UserMemId = 0
    'Utilis� pour r�f�rencer un �l�ment dans la collection
    'vntIndexKey contient l'index ou la cl� dans la collection
    'Il est donc d�clar� en tant que Variant
    'Syntaxe: Set toto = x.Item(xyz) ou Set toto = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'utilis�e pour lire le nombre d'�l�ments dans la
    'collection. Syntaxe: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'utilis�e lors de la suppression d'un �l�ment de la collection
    'vntIndexKey contient l'index ou la cl�, il est donc
    'd�clar� en tant que Variant
    'Syntaxe: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Cette propri�t� vous permet d'�num�rer
    'cette collection avec la syntaxe For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Cr�e la collection lorsque cette classe est cr��e
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'D�truit la collection lorsque cette classe est d�truite
    Set mCol = Nothing
End Sub

