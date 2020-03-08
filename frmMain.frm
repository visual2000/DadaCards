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
   ScaleHeight     =   733
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1021
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

Private Sub Form_Initialize()
    Dim i As Integer
    
    
    For i = 0 To nrOfCards - 1
        
        If Not i = 0 Then
            Load ucCard(i)
        End If
        
        ucCard(i).Left = i * ucCard(i).Width + ucCard(0).Left
        ucCard(i).Visible = True
        
        ucCard(i).Card = i + 1
        
    Next i
End Sub

Private Sub mnuGameQuit_Click()
    Unload frmMain
    Set frmMain = Nothing
End Sub

