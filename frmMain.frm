VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Games"
   ClientHeight    =   3600
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   2565
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   171
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPickup 
      Caption         =   "Play 52 Pickup!"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSolitaire 
      Caption         =   "Play Solitaire!"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPickup_Click()
    Load frmPickup
    frmPickup.Show

End Sub

Private Sub cmdSolitaire_Click()
    Load frmSolitaire
    frmSolitaire.Show
End Sub

Private Sub Form_Load()
    screenWidth = 1024
    screenHeight = 768

    offScreenSet
    SetCardPics
End Sub
