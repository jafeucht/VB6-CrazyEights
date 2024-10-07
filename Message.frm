VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Suit"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   ControlBox      =   0   'False
   Icon            =   "Message.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   390
      Left            =   2175
      TabIndex        =   0
      Top             =   825
      Width           =   1215
   End
   Begin VB.Image imgSuit 
      Height          =   390
      Left            =   150
      Top             =   150
      Width           =   390
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   750
      TabIndex        =   1
      Top             =   150
      Width           =   2640
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MSG_BEGIN = "The suit has just changed to "

Dim EndMsg As String
Dim tSuit As CardTypes

Public Sub ShowNewSuit(SuitType As CardTypes)
    tSuit = SuitType
    Me.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Select Case tSuit
        Case Spades
            EndMsg = "Spades."
        Case Diamonds
            EndMsg = "Diamonds."
        Case Clubs
            EndMsg = "Clubs."
        Case Hearts
            EndMsg = "Hearts."
    End Select
    lblMessage = MSG_BEGIN & EndMsg
    imgSuit.Picture = frmSuit!Suit(tSuit)
End Sub
