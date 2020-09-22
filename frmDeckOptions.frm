VERSION 5.00
Begin VB.Form frmDeckOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose a deck"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   11
      Left            =   6120
      Picture         =   "frmDeckOptions.frx":0000
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   11
      Top             =   1680
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   10
      Left            =   4920
      Picture         =   "frmDeckOptions.frx":5144
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   10
      Top             =   1680
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   9
      Left            =   3720
      Picture         =   "frmDeckOptions.frx":A288
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   9
      Top             =   1680
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   8
      Left            =   2520
      Picture         =   "frmDeckOptions.frx":F3CC
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   8
      Top             =   1680
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   7
      Left            =   1320
      Picture         =   "frmDeckOptions.frx":14510
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   7
      Top             =   1680
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   6
      Left            =   120
      Picture         =   "frmDeckOptions.frx":19654
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   6
      Top             =   1680
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   5
      Left            =   6120
      Picture         =   "frmDeckOptions.frx":1E798
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   5
      Top             =   120
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   4
      Left            =   4920
      Picture         =   "frmDeckOptions.frx":238DC
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   120
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   3
      Left            =   3720
      Picture         =   "frmDeckOptions.frx":28A20
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   3
      Top             =   120
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   2
      Left            =   2520
      Picture         =   "frmDeckOptions.frx":2DB64
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   120
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   1
      Left            =   1320
      Picture         =   "frmDeckOptions.frx":32CA8
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.PictureBox picDeckBack 
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   0
      Left            =   120
      Picture         =   "frmDeckOptions.frx":37DEC
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmDeckOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub picDeckBack_Click(Index As Integer)
    Dim i As Integer
    
    GameDeck.DeckBack = Index + 1
    
    For i = 0 To 6
        Pile(i).DeckBack = Index + 1
    Next
    
    Unload Me
    
End Sub
