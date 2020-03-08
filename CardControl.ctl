VERSION 5.00
Begin VB.UserControl CardControl 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "CardControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()
Public Event DblClick()

Private myCard As Card

Public dragging As Boolean

Public pxWidth As Integer
Public pxHeight As Integer

Public Property Get Card() As Face
  Card = myCard.Face
End Property

Public Property Let Card(ByVal Value As Face)
    myCard.Face = Value
    Call DrawMyself
End Property

Public Property Get faceDown() As Boolean
  faceDown = myCard.faceDown
End Property

Public Property Let faceDown(ByVal Value As Boolean)
    myCard.faceDown = Value
    Call DrawMyself
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Set myCard = New Card
    
    pxWidth = LoadResPicture(101, vbResBitmap).Width
    pxHeight = LoadResPicture(101, vbResBitmap).Height
    
    Width = ScaleX(pxWidth, vbPixels, vbTwips)
    Height = ScaleY(pxHeight, vbPixels, vbTwips)
    ScaleMode = ScaleModeConstants.vbPixels
    MaskColor = &HDFFFFF
    MaskPicture = LoadResPicture(101, vbResBitmap)
    
    dragging = False

    Call DrawMyself
End Sub

Public Sub Flip()
    ' flip the card!
    myCard.faceDown = Not myCard.faceDown
    Call DrawMyself
End Sub

Private Sub DrawMyself()
    If myCard.faceDown Then
        ' Show the generic "card back" picture.
        Picture = LoadResPicture(201, vbResBitmap)
    Else
        Picture = LoadResPicture(myCard.imgResourceId, vbResBitmap)
    End If
End Sub

''' Pass along the events as appropriate!
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragging = True
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragging = False
    ' if haveMoved, don't raise click?
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
