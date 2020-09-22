Attribute VB_Name = "modCards"
Option Explicit

Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Const CARD_WIDTH = 71
Public Const CARD_HEIGHT = 96

Type Rect
    Top As Long
    bottom As Long
    left As Long
    right As Long
End Type

Enum eSuit
    SUIT_CLUBS = 0
    SUIT_DIAMONDS = 1
    SUIT_HEARTS = 2
    SUIT_SPADES = 3
End Enum

Enum eValue
    CARD_ACE = 1
    CARD_TWO
    CARD_THREE
    CARD_FOUR
    CARD_FIVE
    CARD_SIX
    CARD_SEVEN
    CARD_EIGHT
    CARD_NINE
    CARD_TEN
    CARD_JACK
    CARD_QUEEN
    CARD_KING
End Enum

Type tCardPic
    hdc As Long
    pic As StdPicture
End Type

Public screenWidth As Integer
Public screenHeight As Integer

Public offScreenDC As Long
Public offScreenBMP As Long

Public CardPics(0 To 51) As tCardPic
Public ErrorPic As tCardPic
Public DeckBacks(0 To 14) As tCardPic

Sub SetCardPics()
    Dim hTempDC As Long
    Dim i As Integer
    
    hTempDC = GetDC(0)
    
    For i = 0 To 51
    
        Set CardPics(i).pic = New StdPicture
        Set CardPics(i).pic = LoadPicture(App.Path & "\cards\" & i + 1 & ".bmp")
        CardPics(i).hdc = CreateCompatibleDC(hTempDC)
        Call SelectObject(CardPics(i).hdc, CardPics(i).pic.Handle)
    
    Next
    
    For i = 0 To 14
    
        Set DeckBacks(i).pic = New StdPicture
        Set DeckBacks(i).pic = LoadPicture(App.Path & "\Cards\Card backs\" & i + 1 & ".bmp")
        DeckBacks(i).hdc = CreateCompatibleDC(hTempDC)
        Call SelectObject(DeckBacks(i).hdc, DeckBacks(i).pic.Handle)
        
    Next
    
        Set ErrorPic.pic = New StdPicture
        Set ErrorPic.pic = LoadPicture(App.Path & "\Cards\error!.bmp")
        ErrorPic.hdc = CreateCompatibleDC(hTempDC)
        Call SelectObject(ErrorPic.hdc, ErrorPic.pic.Handle)
    
    ReleaseDC 0, hTempDC
End Sub

Public Sub offScreenSet()
  Dim hTempDC As Long
  Dim hOldBMP As Long

  hTempDC = GetDC(0)
  offScreenDC = CreateCompatibleDC(hTempDC)

  offScreenBMP = CreateCompatibleBitmap(hTempDC, screenWidth, screenHeight)
  hOldBMP = SelectObject(offScreenDC, offScreenBMP)
  ReleaseDC 0, hTempDC

End Sub

Function SuitsAreOpposite(Card1 As clsCard, Card2 As clsCard) As Boolean
    If (Card1.Suit = SUIT_CLUBS Or Card1.Suit = SUIT_SPADES) And (Card2.Suit = SUIT_DIAMONDS Or Card2.Suit = SUIT_HEARTS) Then
        SuitsAreOpposite = True
        Exit Function
    End If
    
    If (Card1.Suit = SUIT_DIAMONDS Or Card1.Suit = SUIT_HEARTS) And (Card2.Suit = SUIT_CLUBS Or Card2.Suit = SUIT_SPADES) Then
        SuitsAreOpposite = True
        Exit Function
    End If
End Function

Sub DeleteCardPics()
    Dim i As Integer
    
    For i = 0 To 51
        DeleteDC CardPics(i).hdc
        DeleteObject CardPics(i).pic
    Next
    
    For i = 0 To 14
        DeleteDC DeckBacks(i).hdc
        DeleteObject DeckBacks(i).pic
    Next
    
    DeleteDC ErrorPic.hdc
    DeleteObject ErrorPic.pic
End Sub
