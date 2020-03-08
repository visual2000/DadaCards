VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public faceDown As Boolean

Public Enum Face
    Club_Ace = 1
    Club_Two = 2
    Club_Three = 3
    Club_Four = 4
    Club_Five = 5
    Club_Six = 6
    Club_Seven = 7
    Club_Eight = 8
    Club_Nine = 9
    Club_Ten = 10
    Club_Jack = 11
    Club_Queen = 12
    Club_King = 13

    Diamond_Ace = 14
    Diamond_Two = 15
    Diamond_Three = 16
    Diamond_Four = 17
    Diamond_Five = 18
    Diamond_Six = 19
    Diamond_Seven = 20
    Diamond_Eight = 21
    Diamond_Nine = 22
    Diamond_Ten = 23
    Diamond_Jack = 24
    Diamond_Queen = 25
    Diamond_King = 26

    Heart_Ace = 27
    Heart_Two = 28
    Heart_Three = 29
    Heart_Four = 30
    Heart_Five = 31
    Heart_Six = 32
    Heart_Seven = 33
    Heart_Eight = 34
    Heart_Nine = 35
    Heart_Ten = 36
    Heart_Jack = 37
    Heart_Queen = 38
    Heart_King = 39

    Spade_Ace = 40
    Spade_Two = 41
    Spade_Three = 42
    Spade_Four = 43
    Spade_Five = 44
    Spade_Six = 45
    Spade_Seven = 46
    Spade_Eight = 47
    Spade_Nine = 48
    Spade_Ten = 49
    Spade_Jack = 50
    Spade_Queen = 51
    Spade_King = 52
End Enum

Dim currentFace As Face

Public Property Get Face() As Face
    Card = currentFace
End Property

Public Property Let Face(ByVal newFace As Face)
    currentFace = newFace
End Property

Public Property Get imgResourceId() As Integer
    imgResourceId = currentFace + 100
End Property

Private Sub Class_Initialize()
    faceDown = True
    
    ' Default to first card value.
    Face = Club_Ace
End Sub
