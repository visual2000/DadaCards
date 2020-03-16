VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "Nihilistic card game"
   ClientHeight    =   10995
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar statusScore 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   10740
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   450
      SimpleText      =   "This is test text."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSweep 
      Caption         =   "Sweep into pile"
      Height          =   375
      Left            =   13920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin DadaCards.CardControl ucCard 
      Height          =   1440
      Index           =   0
      Left            =   120
      Top             =   120
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
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuDealing 
      Caption         =   "Card dealing"
      Begin VB.Menu mnuDealRandomly 
         Caption         =   "Randomly"
         Checked         =   -1  'True
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDealFaceup 
         Caption         =   "Face up"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsHide 
         Caption         =   "Hide all"
      End
      Begin VB.Menu mnuToolsFlipAll 
         Caption         =   "Flip all"
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

Dim score As Long

Dim distanceBonus As Long

Dim dragging(0 To nrOfCards - 1) As Boolean
Dim cardOriginLeft As Integer, cardOriginTop As Integer
Dim dragStartX As Integer, dragStartY As Integer

Dim cardWasAtX As Integer, cardWasAtY As Integer

Private Sub cmdSweep_Click()
    Dim i As Integer
    Dim offset As Integer
    offset = 20
    
    For i = 0 To nrOfCards - 1
        ucCard(i).Left = 120 + (Rnd * i) * offset
        ucCard(i).Top = 120 + (Rnd * i) * offset
    Next i
    
    If mnuDealRandomly.Checked Then
        For i = 0 To 1000
            ucCard(Int(Rnd * 52)).ZOrder (0)
        Next i
    End If
    
    score = ((Rnd * 2) - 1) * 10000
    Call showScore
    
End Sub

Private Sub Form_Initialize()
    Randomize
    Dim i As Integer
    For i = 0 To nrOfCards - 1
        dragging(i) = False
        
        If Not i = 0 Then
            Load ucCard(i)
        End If
    Next i
    
    score = 0
    
    Call showScore
    Call PlaceCards
End Sub

Private Sub showScore()
    statusScore.SimpleText = "Score: " & CStr(score)
    statusScore.Panels(1).Text = "Score: " & CStr(score)
    
    'If score < 0 Then
        'statusScore.panels(1).
    
End Sub

Private Sub PlaceCards()
    Dim i As Integer
    Dim maxX As Integer, maxY As Integer
    maxX = frmMain.ScaleWidth - ucCard(0).Width
    maxY = frmMain.ScaleHeight - ucCard(0).Height
    
    For i = 0 To nrOfCards - 1
    
        If mnuDealRandomly.Checked Then
            ucCard(i).Left = Int(Rnd * maxX)
            ucCard(i).Top = Int(Rnd * maxY)
        Else
            ucCard(0).Left = 240
            ucCard(0).Top = 600
            ucCard(i).Left = (i Mod 13) * ucCard(i).Width + ucCard(0).Left
            ucCard(i).Top = Int(i / 13) * ucCard(i).Height + ucCard(0).Top
        End If
                
        ucCard(i).Visible = True
        ucCard(i).faceDown = Not mnuDealFaceup.Checked
        ucCard(i).Card = i + 1
    Next i
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call showScore
End Sub

Private Sub mnuDealRandomly_Click()
    mnuDealRandomly.Checked = Not mnuDealRandomly.Checked
End Sub

Private Sub mnuDealFaceup_Click()
    mnuDealFaceup.Checked = Not mnuDealFaceup.Checked
End Sub

Private Sub mnuGameQuit_Click()
    Unload frmMain
    Set frmMain = Nothing
End Sub

Private Sub mnuGameReset_Click()
    Call PlaceCards
End Sub

Private Sub mnuToolsFlipAll_Click()
    Dim c As CardControl
    For Each c In ucCard
        c.faceDown = Not c.faceDown
    Next c
End Sub

Private Sub mnuToolsHide_Click()
    Dim c As CardControl
    For Each c In ucCard
        c.faceDown = True
    Next c
End Sub



Private Sub ucCard_DblClick(Index As Integer)
    ucCard(Index).Flip
End Sub

Private Sub ucCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ucCard(Index).ZOrder (0)
    dragging(Index) = True
    cardOriginLeft = ucCard(Index).Left
    cardOriginTop = ucCard(Index).Top
    cardWasAtX = ucCard(Index).Left
    cardWasAtY = ucCard(Index).Top
    dragStartX = X
    dragStartY = Y
    distanceBonus = 0
End Sub

Private Sub ucCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragging(Index) Then
        ucCard(Index).Left = cardOriginLeft + X - dragStartX
        ucCard(Index).Top = cardOriginTop + Y - dragStartY
        cardOriginLeft = ucCard(Index).Left
        cardOriginTop = ucCard(Index).Top
        distanceBonus = ((cardWasAtX - ucCard(Index).Left) / 400) ^ 2 + ((cardWasAtY - ucCard(Index).Top) / 400) ^ 2
        'statusScore.SimpleText = "distanceBonus = " & CStr(distanceBonus)
    End If
End Sub

Private Sub ucCard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To nrOfCards - 1
        dragging(i) = False
    Next i
    
    score = score + distanceBonus
    Call showScore
    distanceBonus = 0
End Sub
