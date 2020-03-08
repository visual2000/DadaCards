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
   Begin VB.Image imgCard 
      Height          =   1455
      Left            =   600
      Top             =   840
      Width           =   2415
   End
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

Private myCard As Card

Public pxWidth As Integer
Public pxHeight As Integer

Public Property Get Card() As Face
  Card = myCard.Face
  
End Property

Public Property Let Card(ByVal Value As Face)
    myCard.Face = Value
    Call DrawMyself
    
End Property

Private Sub UserControl_Initialize()
    Set myCard = New Card
    
    pxWidth = LoadResPicture(101, vbResBitmap).Width
    pxHeight = LoadResPicture(101, vbResBitmap).Height
    
    Width = ScaleX(pxWidth, vbPixels, vbTwips)
    Height = ScaleY(pxHeight, vbPixels, vbTwips)
    ScaleMode = ScaleModeConstants.vbPixels
    
    imgCard.Left = 0
    imgCard.Top = 0
    imgCard.Width = pxWidth
    imgCard.Height = pxHeight
    
    Call DrawMyself
End Sub

''' Pass along the events as appropriate!
Private Sub imgCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgCard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgCard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub Flip()
    ' flip the card!
    myCard.faceDown = Not myCard.faceDown
    Call DrawMyself
    
End Sub

Private Sub DrawMyself()
    If myCard.faceDown Then
        ' Show the generic "card back" picture.
        ' imgcard.Picture =
    Else
        imgCard.Picture = LoadResPicture(myCard.imgResourceId, vbResBitmap)
    End If
    
    
End Sub
