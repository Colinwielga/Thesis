VERSION 5.00
Begin VB.Form frmTable 
   BackColor       =   &H0000FF00&
   Caption         =   "BeerBall Table"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to The Main Menu"
      Height          =   2055
      Left            =   9240
      TabIndex        =   2
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Show Me the Table"
      Height          =   1815
      Left            =   9240
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.PictureBox picresults 
      Height          =   12495
      Left            =   600
      ScaleHeight     =   12435
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form shows what a BeerBall table looks like by displaying a picture

Private Sub cmdreturn_Click()
'this subroutine returns the user back to the main menu
frmTable.Hide
frmmain.Show

End Sub

Private Sub cmdshow_Click()
'this subroutine displays a picture of what a Beerball table looks like at the click of a button
    picresults.Picture = LoadPicture(App.Path & "\BeerBallTable.jpg")  'loads picture of the table into picture space
End Sub
