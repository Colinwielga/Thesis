VERSION 5.00
Begin VB.Form frmNotWinner 
   BackColor       =   &H80000007&
   Caption         =   "Not a winner...sorry"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrick 
      BackColor       =   &H8000000D&
      Caption         =   "        Didn't Win??          Still Want A Prize??"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4200
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000D&
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4440
      TabIndex        =   0
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Image imgBummer 
      Height          =   3000
      Left            =   600
      Picture         =   "frmNotWinner.frx":0000
      Top             =   960
      Width           =   3000
   End
End
Attribute VB_Name = "frmNotWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BingoProject
'NotWinner Form
'Missy Ulrich & Beckie Hoyt
'November 3, 2006
'This form allows the user to click on a command button to claim the prize for winning the game of bingo
'The user also has the option of exiting the game through a command button
Option Explicit

Private Sub cmdQuit_Click()
    End
    'Allows the user to exit the program
End Sub

Private Sub cmdTrick_Click()
    MsgBox "Ha! Ha!  Gotcha!  No Prize For You!", , "Claim Your Prize"
    'The above message appears when the user clicks on the command button located on the form
End Sub
