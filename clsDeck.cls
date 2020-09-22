VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const NumCardsInADeck = 52

Public Top As Integer
Public bottom As Integer
Public left As Integer
Public right As Integer

Public DeckBack As Integer
Public NumToDraw As Integer
Public m_NumCards As Integer
Private m_Card() As New clsCard

Property Get Card(index As Integer) As clsCard
    Set Card = m_Card(index)
End Property

Property Let Card(index As Integer, CardValue As clsCard)
    Set m_Card(index) = CardValue
End Property

Private Sub Class_Initialize()

    Dim i As Integer
    Dim currSuit As Integer, currValue As Integer
    
    Randomize
    
    currSuit = 0
    currValue = 1
    
    NumToDraw = 3
    
    m_NumCards = NumCardsInADeck
    DeckBack = Rnd * 11 + 1
    
    ReDim m_Card(m_NumCards)
    
    For i = 0 To m_NumCards - 1
        m_Card(i).Suit = currSuit
        m_Card(i).value = currValue
        
        currValue = currValue + 1
        
        If currValue > 13 Then
            currValue = 1
            currSuit = currSuit + 1
        End If
        
        If currSuit > 4 Then currSuit = 1
    Next
    
End Sub

Public Sub Initialize(DeckSize As Integer, Optional theDeckBack As Integer = -1)
    Dim i As Integer
    Dim currSuit As Integer, currValue As Integer
    
    Randomize

    currSuit = 0
    currValue = 1
    
    m_NumCards = DeckSize
    
    If theDeckBack = -1 Then
        DeckBack = Rnd * 9 + 2
    Else
        DeckBack = theDeckBack
    End If
    
    ReDim m_Card(m_NumCards)
    
    For i = 0 To m_NumCards - 1
        m_Card(i).Suit = currSuit
        m_Card(i).value = currValue
        
        currValue = currValue + 1
        
        If currValue > 13 Then
            currValue = 1
            currSuit = currSuit + 1
        End If
        
        If currSuit > 4 Then currSuit = 1
    Next
End Sub


Public Sub Shuffle()

    Dim i As Integer
    Dim Temp As Integer
    Dim used() As Boolean
    Dim tempCards() As New clsCard
    Dim ShuffleCount As Integer
    
    ReDim used(m_NumCards)
    ReDim tempCards(m_NumCards)
    
    For i = 0 To m_NumCards - 1
        tempCards(i).Suit = m_Card(i).Suit
        tempCards(i).value = m_Card(i).value
    Next
    
    For i = 0 To m_NumCards - 1
        Temp = Int(Rnd * m_NumCards)
        Do While used(Temp) = True
            Temp = Int(Rnd * m_NumCards)
        Loop
        
        m_Card(i).Suit = tempCards(Temp).Suit
        m_Card(i).value = tempCards(Temp).value
        
        used(Temp) = True
    Next
End Sub

Public Sub Draw(X As Integer, Y As Integer, OffsetX As Integer, OffsetY As Integer, Optional theNumToDraw As Integer = -1, Optional FacingUp As Boolean = False, Optional TopCardFacingUp As Boolean = False)
    Dim DrawPosX As Integer, DrawPosY As Integer, i As Integer
    DrawPosX = X
    DrawPosY = Y
    
    If NumToDraw < 0 Then NumToDraw = 0
    If theNumToDraw <> -1 Then NumToDraw = theNumToDraw
    
    Top = Y
    left = X
    right = X + OffsetX * NumToDraw + CARD_WIDTH
    bottom = Y + OffsetY * NumToDraw + CARD_HEIGHT
    
    If m_NumCards = 0 Then
        BitBlt offScreenDC, DrawPosX, DrawPosY, 71, 96, DeckBacks(0).hdc, 0, 0, vbSrcCopy
        Exit Sub
    End If
    
    If NumToDraw > m_NumCards Then NumToDraw = m_NumCards
    
    For i = m_NumCards - NumToDraw To m_NumCards - 1
        If i = m_NumCards - 1 Then
            m_Card(i).FacingUp = TopCardFacingUp
        Else
            m_Card(i).FacingUp = FacingUp
        End If
    
        m_Card(i).Draw DrawPosX, DrawPosY, DeckBack
        
        DrawPosX = DrawPosX + OffsetX
        DrawPosY = DrawPosY + OffsetY
    Next
End Sub

Function DrawCard() As clsCard
    If m_NumCards > 0 Then
        m_NumCards = m_NumCards - 1
        Set DrawCard = m_Card(m_NumCards)
    End If
End Function

Public Sub KillTopCard()
    If m_NumCards > 0 Then
        m_NumCards = m_NumCards - 1
        'ReDim Preserve m_Card(m_NumCards)
    End If
End Sub

Public Sub KillCards(StartIndex As Long, Optional EndIndex As Integer = 1)
    Dim i As Long
    
    For i = StartIndex To EndIndex
        m_Card(i) = m_Card(i + 1)
    Next
    
    m_NumCards = m_NumCards - EndIndex - StartIndex + 1
    
    ReDim Preserve m_Card(m_NumCards)
End Sub

Function GetTopCard() As clsCard
    If m_NumCards > 0 Then
        Set GetTopCard = m_Card(m_NumCards - 1)
    End If
End Function

Sub AddCardToTop(Card As clsCard)
    m_NumCards = m_NumCards + 1
    
    ReDim Preserve m_Card(m_NumCards)
    
    Set m_Card(m_NumCards - 1) = Card
End Sub

Function CheckClicked(ByVal X As Single, ByVal Y As Single) As Boolean
    CheckClicked = (X >= left And X <= right And Y >= Top And Y <= bottom)
End Function
