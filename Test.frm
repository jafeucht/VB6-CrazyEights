VERSION 5.00
Object = "*\ACards.vbp"
Begin VB.Form frmCrazyEights 
   BackColor       =   &H00008000&
   Caption         =   "Crazy Eights"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   Icon            =   "Test.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   2  'CenterScreen
   Begin Cards.Card cdHuman 
      Height          =   1365
      Index           =   0
      Left            =   2775
      TabIndex        =   3
      Top             =   3825
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2408
      FaceMode        =   0
   End
   Begin Cards.Card Deck 
      Height          =   1440
      Left            =   2175
      TabIndex        =   1
      Top             =   2250
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2540
      FaceMode        =   5
   End
   Begin Cards.Card Pile 
      Height          =   1440
      Left            =   3375
      TabIndex        =   0
      Top             =   2250
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2540
   End
   Begin Cards.Card cdComp 
      Height          =   1440
      Index           =   0
      Left            =   2775
      TabIndex        =   2
      Top             =   750
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2540
      FaceMode        =   5
   End
   Begin VB.Label lblHuman 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Human"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   2925
      TabIndex        =   5
      Top             =   5250
      Width           =   765
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   2790
      TabIndex        =   4
      Top             =   450
      Width           =   1035
   End
End
Attribute VB_Name = "frmCrazyEights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CompDraw As Boolean
Dim HumanDraw As Boolean
Dim pComp As PointAPI
Dim pHuman As PointAPI
Dim FaceCard As cCard
Dim NewSuit As CardTypes

Dim CompCards() As cCard
Dim HumanCards() As cCard
Dim DeckCards() As cCard

Const MIN_DIMEN = 6500

Private Sub cdComp_Click(Index As Integer)
    Debug.Print "hello"
End Sub

Private Sub cdHuman_DblClick(Index As Integer)
Dim i As Integer
    If Not IsPlayable(HumanCards(Index)) = None And PlayerTurn = Human Then
        PlaceCard HumanCards(Index)
        For i = Index To UBound(HumanCards) - 1
            HumanCards(i) = HumanCards(i + 1)
            cdHuman(i).SetCardValue cdHuman(i + 1).CardType, cdHuman(i + 1).CardValue
        Next i
        Unload cdHuman(cdHuman.Count - 1)
        If UBound(HumanCards) = 1 Then
            Erase HumanCards
            ShowComputerCards
            frmScore.AddScores GetCardPoints(Human), GetCardPoints(Computer)
            Deal
        Else
            ReDim Preserve HumanCards(1 To UBound(HumanCards) - 1)
            RefreshCards
            PlayerTurn = Computer
            ComputerPlay
        End If
    End If
End Sub

Function GetHighSuit(Player As Players, Disclude As Integer) As Integer
Dim SuitNum As CardTypes, DeckSum As Integer, HighSuitSum
    For SuitNum = Spades To Hearts
        DeckSum = GetSuitSum(Player, SuitNum, Disclude)
        If DeckSum >= HighSuitSum Then
            HighSuitSum = DeckSum
            GetHighSuit = SuitNum
        End If
    Next SuitNum
End Function

Function GetCardPoints(Player As Players) As Integer
Dim SuitNum As CardTypes
    For SuitNum = Spades To Hearts
        GetCardPoints = GetCardPoints + GetSuitSum(Player, SuitNum)
    Next SuitNum
End Function

Sub ComputerPlay()
Dim i As Integer, SuitNum As CardTypes, BestPlay As Integer, HighNum As Integer
Dim HighDeck As Integer, HighDeckSum As Integer, HighSuit As CardTypes, DeckSum As Integer
    Do
        For i = 1 To UBound(CompCards)
            DoEvents
            Select Case IsPlayable(CompCards(i))
                Case CrazyCard
                    HighSuit = GetHighSuit(Computer, i)
                    HighDeck = i
                Case Suit
                    If HighNum = 0 Then
                        HighNum = i
                    ElseIf CompCards(i).cValue > CompCards(HighNum).cValue Then
                        HighNum = i
                    End If
                Case Number
                    DeckSum = GetSuitSum(Computer, CompCards(i).cType, i)
                    If DeckSum >= HighDeckSum Then
                        HighDeckSum = DeckSum
                        HighSuit = CompCards(i).cType
                        HighDeck = i
                    End If
            End Select
        Next i
        If HighNum > 0 Then
            BestPlay = HighNum
        ElseIf HighSuit > 0 Then
            BestPlay = HighDeck
        Else
            Do
                BestPlay = PlayDeck(Computer)
                If BestPlay = -1 Then
                    MsgBox "Computer must draw."
                    PlayerTurn = Human
                    CompDraw = True
                    HumanPlay
                    Exit Sub
                End If
                CompDraw = False
                i = UBound(CompCards)
                If IsPlayable(CompCards(i)) Then
                    BestPlay = i
                    If CompCards(i).cValue = CRAZY_CARD Then
                        HighSuit = GetHighSuit(Computer, i)
                        HighDeck = i
                    End If
                End If
                DoEvents
            Loop Until BestPlay <> 0
        End If
    Loop Until BestPlay > 0
    If CompCards(BestPlay).cValue = CRAZY_CARD Then PlayCrazy HighSuit
    PlaceCard CompCards(BestPlay)
    For i = BestPlay To UBound(CompCards) - 1
        CompCards(i) = CompCards(i + 1)
        cdComp(i).SetCardValue cdComp(i + 1).CardType, cdComp(i + 1).CardValue
    Next i
    Unload cdComp(cdComp.Count - 1)
    If UBound(CompCards) = 1 Then
        Erase CompCards
        frmScore.AddScores GetCardPoints(Human), GetCardPoints(Computer)
        Deal
    Else
        ReDim Preserve CompCards(1 To UBound(CompCards) - 1)
        RefreshCards
        PlayerTurn = Human
        HumanPlay
    End If
End Sub

Sub HumanPlay()
    If Not CheckHumanPlay Then
        HumanDraw = True
        If CompDraw Then
            ShowComputerCards
            frmScore.AddScores GetCardPoints(Human), GetCardPoints(Computer)
            Deal
            Exit Sub
        End If
        MsgBox "You must draw."
        ComputerPlay
    End If
End Sub

Function PlayDeck(PlayerName As Players) As Integer
    On Error Resume Next
    If UBound(DeckCards) = 0 Then
        PlayDeck = -1
        Exit Function
    End If
    Select Case PlayerName
        Case Computer
            ReDim Preserve CompCards(1 To UBound(CompCards) + 1)
            Load cdComp(cdComp.Count)
            CompCards(UBound(CompCards)) = DeckCards(UBound(DeckCards))
        Case Human
            ReDim Preserve HumanCards(1 To UBound(HumanCards) + 1)
            Load cdHuman(cdHuman.Count)
            HumanCards(UBound(HumanCards)) = DeckCards(UBound(DeckCards))
    End Select
    RefreshCards
    If UBound(DeckCards) = 1 Then
        Deck.FaceMode = Base
        Erase DeckCards
    Else
        ReDim Preserve DeckCards(1 To UBound(DeckCards) - 1)
    End If
End Function

Private Sub CopyCardPile(ByRef CopyTo() As cCard, CopyFrom() As cCard)
Dim i As Integer
    On Error Resume Next
    ReDim CopyTo(LBound(CopyFrom) To UBound(CopyFrom))
    For i = LBound(CopyFrom) To UBound(CopyFrom)
        CopyTo(i) = CopyFrom(i)
    Next i
End Sub

Private Function GetSuitSum(Player As Players, SuitType As CardTypes, Optional Disclude As Integer) As Integer
Dim i As Integer, SuitNum As CardTypes, Addition As Integer
Dim CardArray() As cCard
    On Error Resume Next
    If Player = Computer Then
        CopyCardPile CardArray, CompCards
    Else
        CopyCardPile CardArray, HumanCards
    End If
    For i = 1 To UBound(CardArray)
        If i <> Disclude Then
            If CardArray(i).cType = SuitType Then
                Addition = CardArray(i).cValue
                If Addition > 10 Then Addition = 10
                GetSuitSum = GetSuitSum + Addition
            End If
        End If
    Next i
End Function

Private Sub PlaceCard(CardValue As cCard)
    If PlayerTurn = Human And CardValue.cValue = Eight Then
        NewSuit = frmSuit.GetNewSuit
    ElseIf CardValue.cValue <> Eight Then
        NewSuit = 0
    End If
    FaceCard = CardValue
    Pile.SetCardValue CardValue.cType, CardValue.cValue
End Sub

Private Function IsPlayable(CardType As cCard) As PlayCauses
    If CardType.cValue = CRAZY_CARD Then
        IsPlayable = CrazyCard
    ElseIf FaceCard.cType = CardType.cType And NewSuit = 0 Then
        IsPlayable = Suit
    ElseIf CardType.cType = NewSuit Then
        IsPlayable = Suit
    ElseIf FaceCard.cValue = CardType.cValue And NewSuit = 0 Then
        IsPlayable = Number
    End If
End Function

Private Sub Deck_Click()
    PlayDeck Human
    HumanPlay
End Sub

Sub PlayCrazy(nSuit As CardTypes)
    CompDraw = False
    HumanDraw = False
    If PlayerTurn = Computer Then frmMessage.ShowNewSuit nSuit
    NewSuit = nSuit
End Sub

Sub KillCards()
Dim i As Integer
    On Error Resume Next
    For i = cdComp.Count To 1 Step -1
        Unload cdComp.Item(i)
    Next i
    For i = cdHuman.Count To 1 Step -1
        Unload cdHuman.Item(i)
    Next i
End Sub

Public Sub Deal()
Dim i As Integer
    On Error Resume Next
    NewSuit = 0
    Shuffle
    KillCards
    CompDraw = False
    HumanDraw = False
    cdComp(0).FaceMode = Back
    Pile.FaceMode = Back
    Deck.FaceMode = Back
    ReDim CompCards(1 To 7)
    ReDim HumanCards(1 To 7)
    ReDim DeckCards(1 To 37)
    For i = 1 To 7
        CompCards(i) = CardDeck(i)
        HumanCards(i) = CardDeck(i + 7)
        Load cdComp(i)
        Load cdHuman(i)
    Next i
    For i = 1 To 37
        DeckCards(i) = CardDeck(i + 14)
    Next i
    FaceCard = CardDeck(52)
    RefreshCards
    Pile.FaceMode = FaceUp
    If PlayerTurn = Computer Then ComputerPlay
End Sub

Function CheckHumanPlay() As Boolean
Dim i As Integer
    On Error Resume Next
    If UBound(DeckCards) = 0 Then
        For i = 1 To UBound(HumanCards)
            If IsPlayable(HumanCards(i)) Then
                CheckHumanPlay = True
                Exit Function
            End If
        Next i
    Else
        CheckHumanPlay = True
    End If
End Function

Sub RefreshCards()
Dim i As Integer, PileStart As Integer
    If Not Initialized Then Exit Sub
    PileStart = ScaleWidth / 2 - ((UBound(CompCards) - 1) * CARD_SPACE + CARD_WIDTH) / 2
    For i = 1 To UBound(CompCards)
        cdComp(i).Left = PileStart + CARD_SPACE * (i - 1)
        cdComp(i).ZOrder 0
        cdComp(i).Top = pComp.Y
        cdComp(i).Visible = True
        cdComp(i).SetCardValue CompCards(i).cType, CompCards(i).cValue
    Next i
    PileStart = ScaleWidth / 2 - ((UBound(HumanCards) - 1) * CARD_SPACE + CARD_WIDTH) / 2
    For i = 1 To UBound(HumanCards)
        cdHuman(i).Left = PileStart + CARD_SPACE * (i - 1)
        cdHuman(i).ZOrder 0
        cdHuman(i).Top = pHuman.Y
        cdHuman(i).Visible = True
        cdHuman(i).SetCardValue HumanCards(i).cType, HumanCards(i).cValue
    Next i
    Pile.SetCardValue FaceCard.cType, FaceCard.cValue
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static K1 As String, K2 As String, K3 As String
    K3 = K2
    K2 = K1
    K1 = UCase$(Chr$(KeyCode))
    ' Cheat code.
    If K3 = "D" And K2 = "T" And K1 = "X" Then
        ShowComputerCards
    End If
End Sub

Sub ShowComputerCards()
Dim i As Integer
    For i = 0 To cdComp.Count - 1
        cdComp(i).FaceMode = FaceUp
    Next i
End Sub

Private Sub Form_Resize()
    Initialized = True
    If Width < MIN_DIMEN Then Width = MIN_DIMEN
    If Height < MIN_DIMEN Then Height = MIN_DIMEN
    Deck.Top = ScaleHeight / 2 - CARD_HEIGHT / 2
    Pile.Top = Deck.Top
    Deck.Left = (ScaleWidth / 2) - 0.01 * ScaleWidth - CARD_WIDTH
    Pile.Left = ScaleWidth - Deck.Left - CARD_WIDTH
    pComp.Y = Deck.Top / 2 - CARD_HEIGHT / 2
    pHuman.Y = ScaleHeight - pComp.Y - CARD_HEIGHT
    pComp.X = ScaleWidth / 2 - CARD_WIDTH / 2
    pHuman.X = pComp.X
    lblComp.Top = pComp.Y - lblComp.Height - 2
    lblHuman.Top = pHuman.Y + CARD_HEIGHT + 2
    lblComp.Left = ScaleWidth / 2 - lblComp.Width / 2
    lblHuman.Left = ScaleWidth / 2 - lblHuman.Width / 2
    cdComp(0).Move pComp.X, pComp.Y
    cdHuman(0).Move pHuman.X, pHuman.Y
    RefreshCards
End Sub
