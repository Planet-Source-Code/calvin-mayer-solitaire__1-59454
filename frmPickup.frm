VERSION 5.00
Begin VB.Form frmPickup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "52 Pickup!!!"
   ClientHeight    =   7365
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7500
   Icon            =   "frmPickup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Pickup_Timer 
      Interval        =   1
      Left            =   2760
      Top             =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmPickup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tCard
    posX As Long
    posY As Long
    card As New clsCard
End Type

Dim time As Long

Dim numCards As Integer
Dim cardsLeft As Integer

Dim started As Boolean

Dim Deck As New clsDeck
Dim card() As tCard

Private Sub Form_Load()
    Pickup_NewGame
End Sub

Private Sub Pickup_Render()
    Dim i As Integer
    
    Rectangle offScreenDC, 0, 0, screenWidth, screenHeight
    
    For i = numCards - 1 To 0 Step -1
        card(i).card.Draw card(i).posX, card(i).posY
    Next
End Sub

Private Sub Form_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, j As Integer, t As Double
    
    If started = False Then started = True
    
    For i = 0 To cardsLeft - 1
        If card(i).card.Clicked(X, Y) Then
            Deck.AddCardToTop card(i).card
            For j = i To numCards - 1
                Set card(j).card = card(j + 1).card
                card(j).posX = card(j + 1).posX
                card(j).posY = card(j + 1).posY
            Next
            cardsLeft = cardsLeft - 1
            Exit For
        End If
    Next
    
    t = time / 100
    
    If cardsLeft = 0 Then
        MsgBox "It took you " & t & " seconds!", , "Game Over"
        Pickup_NewGame
    End If
End Sub


Sub Pickup_NewGame()
    Dim i As Integer
    numCards = 52
    cardsLeft = numCards
    
    time = 0
    started = False
    
    ReDim card(numCards)
    
    Deck.Initialize numCards
    
    For i = 0 To numCards - 1
        Set card(i).card = Deck.DrawCard
        card(i).posX = Int(Rnd * (Me.ScaleWidth - CARD_WIDTH))
        card(i).posY = Int(Rnd * (Me.ScaleHeight - CARD_HEIGHT))
        card(i).card.FacingUp = True
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteCardPics

    Load frmMain
    frmMain.Show
End Sub

Private Sub Pickup_Timer_Timer()
    Pickup_Render
    
    If started Then time = time + 1
    Label1.Caption = time / 100
    
    BitBlt Me.hdc, 0, 0, screenWidth, screenHeight, offScreenDC, 0, 0, vbSrcCopy
End Sub
