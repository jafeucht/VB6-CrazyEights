VERSION 5.00
Begin VB.UserControl Card 
   BackColor       =   &H00008000&
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   ClipControls    =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   1065
   Begin VB.Shape shpSelected 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   5
      Height          =   1440
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CardTypes
    Spades = 1
    Diamonds
    Clubs
    Hearts
End Enum

Public Enum CardValues
    Ace = 1
    Two
    Three
    Four
    Five
    Six
    Seven
    Eight
    Nine
    Ten
    Jack
    Queen
    King
End Enum

Public Enum FaceModes
    FaceUp
    Blank
    Base
    Circled
    Crossed
    Back
End Enum

Dim cValue As CardValues
Dim cType As CardTypes
Dim cSelected As Boolean
Dim cFaceMode As FaceModes
Dim BackPic As Integer

Event SelectCard()
Event DeselectCard()
Event Click()
Event DblClick()

Property Get FaceMode() As FaceModes
    FaceMode = cFaceMode
End Property

Property Let FaceMode(NewValue As FaceModes)
    cFaceMode = NewValue
    RefreshCard
    PropertyChanged "FaceMode"
End Property

Property Get Selected() As Boolean
    Selected = cSelected
End Property

Property Let Selected(NewValue As Boolean)
    If cSelected = NewValue Then Exit Property
    cSelected = NewValue
    If NewValue Then RaiseEvent SelectCard Else: RaiseEvent DeselectCard
    RefreshCard
    PropertyChanged "Selected"
End Property

Sub RefreshCard()
    shpSelected.Visible = cSelected
    If cFaceMode = FaceUp Then
        GetCard
    Else
        GetFace
    End If
End Sub

Property Get CardValue() As CardValues
    CardValue = cValue
End Property

Property Let CardValue(NewValue As CardValues)
    cValue = NewValue
    RefreshCard
    PropertyChanged "CardValue"
End Property

Property Get CardType() As CardTypes
    CardType = cType
End Property

Property Let CardType(NewValue As CardTypes)
    cType = NewValue
    RefreshCard
    PropertyChanged "CardType"
End Property

Private Sub imgCard_Click()
    RaiseEvent Click
End Sub

Private Sub imgCard_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    cValue = Ace
    cType = Spades
    UserControl.Width = 1065
    UserControl.Height = 1440
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    cFaceMode = PropBag.ReadProperty("FaceMode", Blank)
    cSelected = PropBag.ReadProperty("Selected", False)
    cValue = PropBag.ReadProperty("CardValue", Ace)
    cType = PropBag.ReadProperty("CardType", Spades)
    RefreshCard
End Sub

Private Sub UserControl_Resize()
    imgCard.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Sub GetCard()
    If cValue = 0 Then Exit Sub
    Set imgCard.Picture = LoadResPicture(100 + 4 * (cValue - 1) + cType, 0)
End Sub

Public Sub SetCardValue(cdType As CardTypes, cdValue As CardValues)
    cType = cdType
    cValue = cdValue
    RefreshCard
End Sub

Function GetFace()
    If cFaceMode = FaceUp Then Exit Function
    If cFaceMode = Back Then cFaceMode = Back + BackPic
    Set imgCard.Picture = LoadResPicture(152 + cFaceMode, 0)
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "FaceMode", cFaceMode, Blank
    PropBag.WriteProperty "Selected", cSelected, False
    PropBag.WriteProperty "CardValue", cValue, Ace
    PropBag.WriteProperty "CardType", cType, Spades
End Sub
