Attribute VB_Name = "modSolitaireGame"
'Copyright 2005 Calvin Mayer

Option Explicit

Public numCardsToDraw As Integer
Public Score As Long
Public keepingScore As Boolean
Public timingTheGame As Boolean
Public time As Long

Dim bla4 As Boolean

Type tSelectedCard
    card() As New clsCard
    PosFrom As Integer
End Type

Type tPileCard
    numCards As Integer
    card() As New clsCard
End Type

Public GameDeck As New clsDeck
Public DispDeck As New clsDeck
Public GoalDeck(0 To 3) As New clsDeck
Public Pile(0 To 6) As New clsDeck

Public CardPile(0 To 6) As tPileCard
Public SelectedCard As tSelectedCard
Public NumSelected As Integer
Public CardIsSelected As Boolean

Sub Solitaire_New_Game()
    Dim i As Integer, j As Integer
    
    
    Randomize
    
    CardIsSelected = False
    NumSelected = 1
    ReDim SelectedCard.card(NumSelected)
    
    Score = 0
    time = 0
    frmSolitaire.lblTime.Caption = ""
    If keepingScore Then
        frmSolitaire.lblScore.Caption = "Score: 0"
    Else
        frmSolitaire.lblScore.Caption = ""
    End If
    SelectedCard.card(NumSelected).FacingUp = True
    SelectedCard.PosFrom = -1

    GameDeck.Initialize 52
    GameDeck.Shuffle
    
    DispDeck.Initialize 0, GameDeck.DeckBack
    
    For i = 0 To 6
        Pile(i).Initialize 0, GameDeck.DeckBack
    Next

    For i = 0 To 6
        For j = 0 To i
            Pile(i).AddCardToTop GameDeck.DrawCard
        Next
        CardPile(i).numCards = 1
        ReDim CardPile(i).card(CardPile(i).numCards)
        Set CardPile(i).card(0) = Pile(i).DrawCard
    Next
    
    For i = 0 To 3
        GoalDeck(i).Initialize 0, GameDeck.DeckBack
    Next
    
    
End Sub

Sub Solitaire_CheckMouseDown(ByVal X As Single, ByVal Y As Single)
    Dim i As Integer, j As Integer, k As Integer
    Dim done As Boolean, bla As Boolean, bla2 As Boolean, bla3 As Boolean
    
    bla4 = False
    
    If GameDeck.CheckClicked(X, Y) Then
        If CardIsSelected = False Then
            If GameDeck.m_NumCards = 0 Then
                For i = 0 To DispDeck.m_NumCards - 1
                    GameDeck.AddCardToTop DispDeck.DrawCard
                    done = True
                    AddToScore -Score
                Next
            End If
            If done = False Then
                For i = 0 To numCardsToDraw - 1
                    If GameDeck.m_NumCards = 0 Then Exit For
                    DispDeck.AddCardToTop GameDeck.DrawCard
                Next
                DispDeck.NumToDraw = 3
            End If
        End If
    End If
    
    
    If DispDeck.CheckClicked(X, Y) Then
        If CardIsSelected = False Then
            If DispDeck.m_NumCards > 0 Then
                NumSelected = 1
                Set SelectedCard.card(0) = DispDeck.DrawCard
                If DispDeck.NumToDraw > 1 Then DispDeck.NumToDraw = DispDeck.NumToDraw - 1
                SelectedCard.PosFrom = 0
                CardIsSelected = True
            End If
        Else
            ReturnSelectedCard
        End If
    End If

    
    
    For i = 0 To 3
        If GoalDeck(i).CheckClicked(X, Y) Then
            If CardIsSelected = True Then
                If SelectedCard.PosFrom = 0 Or SelectedCard.PosFrom > 4 Then AddToScore 10
                AddCardsToGoalDecks i

            Else
                If GoalDeck(i).m_NumCards > 0 Then
                    NumSelected = 1
                    ReDim SelectedCard.card(NumSelected)
                    SelectedCard.PosFrom = i + 1
                    Set SelectedCard.card(0) = GoalDeck(i).DrawCard
                    CardIsSelected = True

                End If
            End If
        End If
    Next
    
    done = False
    
    'can't remember what these vars are for :S I couldn't think of meaningful
    'names at the time
    bla = True
    bla2 = False
    bla3 = False
    
    
    For i = 0 To 6
        If Pile(i).CheckClicked(X, Y) Then
            If CardPile(i).numCards = 0 Then
                If CardIsSelected = False Then
                    If Pile(i).m_NumCards > 0 Then
                        CardPile(i).numCards = 1
                        ReDim CardPile(i).card(CardPile(i).numCards)
                        Set CardPile(i).card(0) = Pile(i).DrawCard
                        CardPile(i).card(0).FacingUp = True
                        done = True
                    End If
                Else
                    If SelectedCard.card(0).value = CARD_KING And Pile(i).m_NumCards = 0 And done = False Then
                        If SelectedCard.PosFrom = 0 Then AddToScore 5
                        If SelectedCard.PosFrom >= 1 And SelectedCard.PosFrom <= 4 Then AddToScore -15
                        AddCardsToPiles i
                        bla3 = True
                    Else
                        bla2 = True
                        ReturnSelectedCard
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    
    bla = False
    
    For i = 0 To 6
        For j = CardPile(i).numCards - 1 To 0 Step -1
            If CardPile(i).card(j).Clicked(X, Y) Then
                If CardIsSelected = False Then
                    If done = False And bla = False And bla2 = False And bla3 = False And bla4 = False Then
                        GetCardsFromPile i, j
                    End If
                    Exit For
                Else
                    If CardPile(i).numCards > 0 Then
                        If (SuitsAreOpposite(SelectedCard.card(0), CardPile(i).card(CardPile(i).numCards - 1))) And (SelectedCard.card(0).value = CardPile(i).card(CardPile(i).numCards - 1).value - 1) Then
                            If SelectedCard.PosFrom = 0 Then
                                AddToScore 5
                            ElseIf SelectedCard.PosFrom >= 1 And SelectedCard.PosFrom <= 4 Then
                                AddToScore -15
                            End If
                            AddCardsToPiles i
                            Exit For
                        Else
                            ReturnSelectedCard
                            bla = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    Next
    
    If GoalDeck(0).m_NumCards = 14 And GoalDeck(1).m_NumCards = 14 And GoalDeck(2).m_NumCards = 14 And GoalDeck(3).m_NumCards = 14 Then
        MsgBox "Hurrah, you won.", , "Solitaire!"
        Solitaire_New_Game
    End If

    
    'For i = 0 To 6
    '    If Pile(i).CheckClicked(X, Y) Then
    '        If CardIsSelected = True Then
    '            If Pile(i).m_NumCards = 0 And CardPile(i).numCards = 0 Then
    '                If SelectedCard.Card.value = CARD_KING Then
    '                    AddCardsToPiles i
    '                    CardIsSelected = False
    '                End If
    '            Else
    '                'ReturnSelectedCard
    '            End If
    '        Else
    '            If CardPile(i).numCards = 0 Then
    '                CardPile(i).numCards = CardPile(i).numCards + 1
    '                ReDim Preserve CardPile(i).Card(CardPile(i).numCards)
    '                Set CardPile(i).Card(CardPile(i).numCards - 1) = Pile(i).DrawCard
    '                CardPile(i).Card(CardPile(i).numCards - 1).FacingUp = True
    '                CardIsSelected = False
    '            End If
    '        End If
    '    End If
    'Next
    
End Sub

Sub Solitaire_CheckMouseMove(ByVal X As Single, Y As Single)
    If CardIsSelected Then
        SelectedCard.card(0).left = X - CARD_WIDTH / 2
        SelectedCard.card(0).Top = Y - CARD_HEIGHT / 2
    End If
End Sub

Sub Solitaire_Render(Form As Form)
    Dim i As Integer, j As Integer
    Dim DrawOffsetX As Integer, DrawOffsetY As Integer
    
    DrawOffsetX = 2
    DrawOffsetY = 15
    
    frmSolitaire.Cls
    
    Rectangle offScreenDC, 0, 0, screenWidth, screenHeight
    
    GameDeck.Draw 10, 10, 2, 2, 5
    DispDeck.Draw 100, 10, 15, 0, , True, True
    
    For i = 0 To 3
        GoalDeck(i).Draw 280 + 90 * i, 10, 1, 2, 5, True, True
    Next
    
    For i = 0 To 6
        Pile(i).Draw 10 + 90 * i, 130, 2, 3, 7
    Next
    
    For i = 0 To 6
        For j = 0 To CardPile(i).numCards - 1
            CardPile(i).card(j).Draw (Pile(i).left + 2 * Pile(i).m_NumCards) + DrawOffsetX * j, (Pile(i).Top + 3 * Pile(i).m_NumCards) + DrawOffsetY * j
        Next
    Next
    
    If CardIsSelected Then
        For i = 0 To NumSelected - 1
            SelectedCard.card(i).Draw SelectedCard.card(0).left + DrawOffsetX * i, SelectedCard.card(0).Top + DrawOffsetY * i
        Next
    End If
    
    BitBlt Form.hdc, 0, 0, screenWidth, screenHeight, offScreenDC, 0, 0, vbSrcCopy
    
    frmSolitaire.Refresh
End Sub

Private Sub ReturnSelectedCard()
    Dim Index As Integer

    If SelectedCard.PosFrom = 0 Then
        DispDeck.AddCardToTop SelectedCard.card(0)
        DispDeck.NumToDraw = DispDeck.NumToDraw + 1
        CardIsSelected = False
    ElseIf SelectedCard.PosFrom >= 1 And SelectedCard.PosFrom <= 4 Then
        GoalDeck(SelectedCard.PosFrom - 1).AddCardToTop SelectedCard.card(0)
        CardIsSelected = False
    ElseIf SelectedCard.PosFrom >= 5 And SelectedCard.PosFrom <= 11 Then
        Index = SelectedCard.PosFrom - 5
        AddCardsToPiles Index
        
    End If
    
    CardIsSelected = False
End Sub

Private Sub AddCardsToGoalDecks(Index As Integer)
    Dim ItsAllGood As Boolean
    
    If NumSelected = 1 Then
        If GoalDeck(Index).m_NumCards = 0 Then
            If SelectedCard.card(0).value = CARD_ACE Then
                ItsAllGood = True
                GoalDeck(Index).AddCardToTop SelectedCard.card(0)
                CardIsSelected = False
                
            End If
        Else
            If SelectedCard.card(0).value = GoalDeck(Index).GetTopCard.value + 1 Then
                If SelectedCard.card(0).Suit = GoalDeck(Index).GetTopCard.Suit Then
                    ItsAllGood = True
                    GoalDeck(Index).AddCardToTop SelectedCard.card(0)
                    CardIsSelected = False
                End If
            End If
        End If
    Else
        ReturnSelectedCard
        ItsAllGood = True
        bla4 = True
    End If
    
    If ItsAllGood = False Then
        ReturnSelectedCard
        NumSelected = 0
    Else
        SelectedCard.PosFrom = Index + 1
    End If
    
    CardIsSelected = False
End Sub

Private Sub AddCardsToPiles(Index As Integer)
    Dim i As Integer
    For i = 0 To NumSelected - 1
        CardPile(Index).numCards = CardPile(Index).numCards + 1
        ReDim Preserve CardPile(Index).card(CardPile(Index).numCards)
        Set CardPile(Index).card(CardPile(Index).numCards - 1) = SelectedCard.card(i)

    Next
    
    CardIsSelected = False
End Sub

Private Sub GetCardsFromPile(pileIndex As Integer, cardIndex As Integer)
    Dim i As Integer, curIndex As Integer
    
    curIndex = 0
    
    SelectedCard.PosFrom = pileIndex + 5
    NumSelected = CardPile(pileIndex).numCards - cardIndex
    ReDim SelectedCard.card(NumSelected)
    
    For i = cardIndex To CardPile(pileIndex).numCards - 1
        Set SelectedCard.card(curIndex) = CardPile(pileIndex).card(i)
        curIndex = curIndex + 1
    Next
    
    CardPile(pileIndex).numCards = CardPile(pileIndex).numCards - NumSelected
    
    CardIsSelected = True
    
End Sub

Private Sub AddToScore(ByVal amount As Integer)
    If keepingScore Then
        Score = Score + amount
        If Score < 0 Then Score = 0
        frmSolitaire.lblScore.Caption = "Score: " & Score
    End If
End Sub
