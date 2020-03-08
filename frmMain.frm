VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "Nihilistic card game"
   ClientHeight    =   10995
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15315
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   15315
   StartUpPosition =   3  'Windows Default
   Begin DadaCards.CardControl ucCard 
      Height          =   1440
      Index           =   0
      Left            =   2160
      Top             =   1080
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   2540
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameReset 
         Caption         =   "Reset"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuGameQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const nrOfCards = 52

Dim dragging(0 To nrOfCards - 1) As Boolean
Dim cardOriginLeft As Integer, cardOriginTop As Integer
Dim dragStartX As Integer, dragStartY As Integer

Private Sub Form_Initialize()
    Randomize
    Dim i As Integer
    For i = 0 To nrOfCards - 1
        dragging(i) = False
        
        If Not i = 0 Then
            Load ucCard(i)
        End If
    Next i
    
    Call PlaceCards
End Sub

Private Sub PlaceCards()
    Dim i As Integer
    Dim maxX As Integer, maxY As Integer
    maxX = frmMain.ScaleWidth - ucCard(0).Width
    maxY = frmMain.ScaleHeight - ucCard(0).Height
    
    For i = 0 To nrOfCards - 1
        'ucCard(i).Left = i * ucCard(i).Width + ucCard(0).Left
        ucCard(i).Left = Int(Rnd * maxX)
        ucCard(i).Top = Int(Rnd * maxY)
        
        ucCard(i).Visible = True
        
        ucCard(i).Card = i + 1
    Next i
End Sub

Private Sub mnuGameQuit_Click()
    Unload frmMain
    Set frmMain = Nothing
End Sub

Private Sub ucCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragging(Index) = True
    cardOriginLeft = ucCard(Index).Left
    cardOriginTop = ucCard(Index).Top
    dragStartX = X
    dragStartY = Y
    
End Sub

Private Sub ucCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragging(Index) Then
        ucCard(Index).Left = cardOriginLeft + X - dragStartX
        ucCard(Index).Top = cardOriginTop + Y - dragStartY
        cardOriginLeft = ucCard(Index).Left
        cardOriginTop = ucCard(Index).Top
    End If
End Sub

Private Sub ucCard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To nrOfCards - 1
        dragging(i) = False
    Next i
End Sub
