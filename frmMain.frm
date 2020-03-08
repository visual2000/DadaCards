VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "Dadaist card game"
   ClientHeight    =   10995
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   15315
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   480
      Top             =   600
      Width           =   2535
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
Dim dragging As Boolean
Dim originalLeft As Integer, originalTop As Integer
Dim startDragX As Integer, startDragY As Integer


Private Sub Form_Load()
    Image1.Picture = LoadResPicture(101, vbResBitmap)
    dragging = False
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print ("mouse move " + CStr(X) + "x" + CStr(Y))
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '' todo only react on mouse-button-1
    'Debug.Print "mouse down"
    dragging = True
    originalLeft = Image1.Left
    originalTop = Image1.Top
    startDragX = X
    startDragY = Y
        

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print ("mouse move " + CStr(X) + "x" + CStr(Y))
    
    If dragging Then
        Image1.Left = originalLeft + X - startDragX
        Image1.Top = originalTop + Y - startDragY
        originalLeft = Image1.Left
        originalTop = Image1.Top
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print ("mouse up")
    dragging = False
    
End Sub


Private Sub mnuGameQuit_Click()
    Unload frmMain
    Set frmMain = Nothing
End Sub
