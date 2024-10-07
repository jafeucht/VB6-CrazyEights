Attribute VB_Name = "mdCrazyEights"
Option Explicit

Enum Players
    Human
    Computer
End Enum

Enum PlayCauses
    None
    Suit
    Number
    CrazyCard
End Enum

Public Initialized As Boolean
Public PlayerTurn As Players

Type PointAPI
    X As Integer
    Y As Integer
End Type

Public Type cCard
    cType As CardTypes
    cValue As CardValues
End Type

Global Const CARD_WIDTH = 71
Global Const CARD_HEIGHT = 91
Global Const CARD_SPACE = 15
Global Const CRAZY_CARD = 8
Global Const CARDS_IN_DECK = 52

Public CardDeck(1 To 52) As cCard

Public Sub ResetDeck()
Dim i As Integer
    For i = 1 To 52
        CardDeck(i).cType = (i - 1) \ 13 + 1
        CardDeck(i).cValue = (i - 1) Mod 13 + 1
    Next i
End Sub

Public Function Card(CardValue As CardValues, CardType As CardTypes) As cCard
    Card.cType = CardType
    Card.cValue = CardValue
End Function

Public Sub Shuffle()
Dim cT As CardTypes, cV As CardValues, rNum As Integer
    Randomize Timer
    Erase CardDeck
    For cT = Spades To Hearts
        For cV = Ace To King
            Do
                rNum = Int(Rnd * 52) + 1
            Loop Until CardDeck(rNum).cType = 0
            CardDeck(rNum).cType = cT
            CardDeck(rNum).cValue = cV
        Next cV
    Next cT
End Sub

Sub Main()
    ResetDeck
    frmCrazyEights.Deal
    frmCrazyEights.Show
End Sub
