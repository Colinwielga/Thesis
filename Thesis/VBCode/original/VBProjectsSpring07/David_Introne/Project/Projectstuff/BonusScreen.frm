VERSION 5.00
Begin VB.Form BonusScreen 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Nice Job! click Anywhere to recieve your free 7 points!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   5445
      Left            =   6000
      Picture         =   "BonusScreen.frx":0000
      Top             =   0
      Width           =   6000
   End
   Begin VB.Image Image1 
      Height          =   4785
      Left            =   0
      Picture         =   "BonusScreen.frx":6950
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "BonusScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Image1_Click()
    VetVisit.Show ' Goes back
    Score = Score + 7 'adds to score
    BonusScreen.Hide
End Sub

Private Sub Image2_Click()
    VetVisit.Show ' Goes back
    BonusScreen.Hide
    Score = Score + 7 'adds to score
End Sub
