VERSION 5.00
Begin VB.Form frmSolitaireOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox ckTime 
      Caption         =   "Timed Game"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Frame frmScore 
      Caption         =   "Scoring"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optScore 
         Caption         =   "None"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optScore 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Draw"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton optDraw 
         Caption         =   "1 Card (wimp!)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optDraw 
         Caption         =   "3 Cards (real solitaire)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmSolitaireOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()

    'these vars check if any real changes were made so we only start a new
    'game if changes were made
    Dim threeCardsDrawn As Boolean
    Dim normalScoring As Boolean
    Dim timingGame As Boolean
    Dim startNewGame As Boolean
    
    threeCardsDrawn = (numCardsToDraw = 3)
    normalScoring = keepingScore
    timingGame = timingTheGame
    
    If optDraw(0).value = True Then
        If threeCardsDrawn = False Then startNewGame = True
        numCardsToDraw = 3
    Else
        If threeCardsDrawn = True Then startNewGame = True
        numCardsToDraw = 1
    End If
    
    If optScore(0).value = True Then
        If normalScoring = False Then startNewGame = True
        keepingScore = True
    Else
        If normalScoring = True Then startNewGame = True
        keepingScore = False
    End If
    
    If ckTime.value = 1 Then
        If timingGame = False Then startNewGame = True
        timingTheGame = True
    Else
        If timingGame = True Then startNewGame = True
        timingTheGame = False
    End If

    
    If startNewGame Then Solitaire_New_Game
    
    Unload Me
End Sub

Private Sub Form_Load()
    If numCardsToDraw = 3 Then
        optDraw(0).value = True
    Else
        optDraw(1).value = True
    End If
    
    If keepingScore Then
        optScore(0).value = True
    Else
        optScore(1).value = True
    End If
    
    If timingTheGame Then
        ckTime.value = 1
    Else
        ckTime.value = 0
    End If
End Sub
