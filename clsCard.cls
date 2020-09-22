VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Top As Integer
Public bottom As Integer
Public left As Integer
Public right As Integer
Private cardIndex As Integer
Private m_Suit As eSuit
Private m_Value As eValue
Private m_FacingUp As Boolean

Property Get Card_Index()
    Card_Index = cardIndex
End Property

Property Get Suit() As eSuit
    Suit = m_Suit
End Property

Property Let Suit(value As eSuit)
    m_Suit = value
    cardIndex = m_Suit * 13 + m_Value - 1
End Property

Property Get value() As eValue
    value = m_Value
End Property

Property Let value(CardValue As eValue)
    m_Value = CardValue
    cardIndex = m_Suit * 13 + m_Value - 1
End Property

Property Get FacingUp() As Boolean
    FacingUp = m_FacingUp
End Property

Property Let FacingUp(value As Boolean)
    m_FacingUp = value
End Property

Public Sub Draw(ByVal X As Integer, ByVal Y As Integer, Optional BackIndex As Integer = 0)
    Top = Y
    left = X
    right = X + CARD_WIDTH
    bottom = Y + CARD_HEIGHT
    
    If cardIndex <> -1 Then
    
        If (m_Value = 0) And (m_Suit = 0) Then
            BitBlt offScreenDC, X, Y, CARD_WIDTH, CARD_HEIGHT, ErrorPic.hdc, 0, 0, vbSrcCopy
        Else
            If m_FacingUp Then
                BitBlt offScreenDC, X, Y, CARD_WIDTH, CARD_HEIGHT, CardPics(cardIndex).hdc, 0, 0, vbSrcCopy
            Else
                BitBlt offScreenDC, X, Y, CARD_WIDTH, CARD_HEIGHT, DeckBacks(BackIndex).hdc, 0, 0, vbSrcCopy
            End If
        
            
        End If
    
    End If
End Sub

Private Sub Class_Initialize()
    m_FacingUp = True
End Sub

Public Function Clicked(ByVal X As Single, ByVal Y As Single) As Boolean
    Clicked = False
    If cardIndex <> -1 Then
        If X >= left And X <= right And Y >= Top And Y <= bottom Then
            Clicked = True
        End If
    End If
End Function
