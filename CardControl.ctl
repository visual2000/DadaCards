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

Public Property Get Card() As Face
  Card = myCard.Face
  
End Property

Public Property Let Card(ByVal Value As Face)
    myCard.Face = Value
    imgCard.Picture = LoadResPicture(myCard.imgResourceId, vbResBitmap)
  
End Property

Private Sub UserControl_Initialize()
    Set myCard = New Card
    
    Width = ScaleX(72, vbPixels, vbTwips)
    Height = ScaleY(96, vbPixels, vbTwips)
    ScaleMode = ScaleModeConstants.vbPixels
    
    imgCard.Left = 0
    imgCard.Top = 0
    imgCard.Width = 72
    imgCard.Height = 96
    
    imgCard.Picture = LoadResPicture(101, vbResBitmap)
    
    
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

