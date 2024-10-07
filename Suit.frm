VERSION 5.00
Begin VB.Form frmSuit 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose a new Suit"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2340
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Suit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   2340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblHearts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hearts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   975
      TabIndex        =   3
      Top             =   1725
      Width           =   600
   End
   Begin VB.Label lblClubs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clubs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   975
      TabIndex        =   2
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label lblDiamonds 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diamonds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   975
      TabIndex        =   1
      Top             =   675
      Width           =   930
   End
   Begin VB.Label lblSpades 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   975
      TabIndex        =   0
      Top             =   150
      Width           =   720
   End
   Begin VB.Image Suit 
      Height          =   330
      Index           =   4
      Left            =   225
      Picture         =   "Suit.frx":000C
      Top             =   1725
      Width           =   360
   End
   Begin VB.Image Suit 
      Height          =   330
      Index           =   3
      Left            =   225
      Picture         =   "Suit.frx":0196
      Top             =   1200
      Width           =   360
   End
   Begin VB.Image Suit 
      Height          =   330
      Index           =   1
      Left            =   225
      Picture         =   "Suit.frx":0320
      Top             =   150
      Width           =   360
   End
   Begin VB.Image Suit 
      Height          =   330
      Index           =   2
      Left            =   225
      Picture         =   "Suit.frx":04AA
      Top             =   675
      Width           =   360
   End
End
Attribute VB_Name = "frmSuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NewSuit As CardTypes

Public Function GetNewSuit() As CardTypes
    Show vbModal
    GetNewSuit = NewSuit
End Function

Private Sub Suit_Click(Index As Integer)
    NewSuit = Index
    Unload Me
End Sub
