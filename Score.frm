VERSION 5.00
Begin VB.Form frmScore 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Score"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   ControlBox      =   0   'False
   Icon            =   "Score.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Okay"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   825
      TabIndex        =   0
      Top             =   1800
      Width           =   1890
   End
   Begin VB.Label hScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1950
      TabIndex        =   5
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label cScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1950
      TabIndex        =   4
      Top             =   975
      Width           =   915
   End
   Begin VB.Label lblHuman 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Human:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   675
      TabIndex        =   3
      Top             =   1425
      Width           =   915
   End
   Begin VB.Label lblComputer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Computer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   675
      TabIndex        =   2
      Top             =   1050
      Width           =   915
   End
   Begin VB.Label lblWinner 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   3165
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CompSc As Integer
Public HumanSc As Integer

Const MSG_COMPUTERWIN = "You let the computer win, you loser. The computer won "
Const MSG_HUMANWIN = "You have just won!!! Great job. You have just earned "
Const MSG_DRAW = "The game has ended in a block. "
Const MSG_DRAW_COMP = "The computer has beat you by "
Const MSG_DRAW_HUMAN = "You beat the computer by"
Const MSG_DRAW_TIE = "Both players earned "
Const MSG_CAPTION = "GAME OVER - "

Public Sub AddScores(ComputerScore As Integer, HumanScore As Integer)
    If ComputerScore = 0 Then
        lblWinner = MSG_HUMANWIN & HumanScore & " points."
        Caption = MSG_CAPTION & "You win."
    ElseIf HumanScore = 0 Then
        Caption = MSG_CAPTION & "Computer wins."
        lblWinner = MSG_COMPUTERWIN & ComputerScore & " points."
    Else
        lblWinner = MSG_DRAW
        If ComputerScore > HumanScore Then
            lblWinner = lblWinner & MSG_DRAW_COMP & (ComputerScore - HumanScore) & " points."
            Caption = MSG_CAPTION & "Computer wins."
        ElseIf HumanScore > ComputerScore Then
            lblWinner = lblWinner & MSG_DRAW_HUMAN & (HumanScore - ComputerScore) & " points."
            Caption = MSG_CAPTION & "You win."
        Else
            lblWinner = lblWinner & MSG_DRAW_TIE & ComputerScore & " points."
            Caption = MSG_CAPTION & "Tie."
        End If
    End If
    CompSc = CompSc + ComputerScore
    HumanSc = HumanSc + HumanScore
    cScore = Format$(CompSc, "000000")
    hScore = Format$(HumanSc, "000000")
    Show vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
