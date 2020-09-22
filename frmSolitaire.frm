VERSION 5.00
Begin VB.Form frmSolitaire 
   AutoRedraw      =   -1  'True
   Caption         =   "Solitaire!"
   ClientHeight    =   7665
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10410
   Icon            =   "frmSolitaire.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   694
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Game_Timer 
      Interval        =   1000
      Left            =   3720
      Top             =   2640
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time: 0"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Score: 0"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuDeck 
         Caption         =   "&Deck Options..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Game"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmSolitaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    numCardsToDraw = 3
    timingTheGame = True
    keepingScore = True
    Solitaire_New_Game
    Solitaire_Render Me
    
End Sub

Private Sub Form_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Solitaire_CheckMouseDown X, Y
    Solitaire_Render Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Solitaire_CheckMouseMove X, Y
    Solitaire_Render Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Solitaire_Render Me
End Sub

Private Sub Form_Terminate()
    DeleteCardPics
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteCardPics
End Sub

Private Sub Game_Timer_Timer()
    If timingTheGame Then
        time = time + 1
        lblTime = "Time: " & time
    End If
End Sub

Private Sub mnuDeck_Click()
    Load frmDeckOptions
    frmDeckOptions.Show

    
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    Solitaire_New_Game
End Sub

Private Sub mnuOptions_Click()
    Load frmSolitaireOptions
    frmSolitaireOptions.Show
End Sub

Private Sub mnuUndo_Click()
    MsgBox "Oh, so you made a mistake and want to undo it, huh? Well guess what bub? There is no undo function in THIS version! HAHAHAHA! Guess you should be more careful next time!", , "Solitaire"
End Sub


