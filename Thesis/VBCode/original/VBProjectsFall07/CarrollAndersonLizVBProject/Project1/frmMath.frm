VERSION 5.00
Begin VB.Form frmMath 
   Caption         =   "Math"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   Picture         =   "frmMath.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdFindThePattern 
      Caption         =   "Find The Pattern"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSudoku 
      Caption         =   "Sudoku"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdMinesweeper 
      Caption         =   "Minesweeper"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "frmMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFindThePattern_Click()

'Takes the user to the Find the Pattern game, and messages them
'how to play

frmMath.Hide
frmFindThePattern.Show
MsgBox ("Find a pattern in each of the following lists of numbers and write the next term in the box.")
End Sub

Private Sub cmdMinesweeper_Click()

'This takes the user to the minesweeper game

frmMath.Hide
frmMinesweeper.Show
End Sub

Private Sub cmdQuit_Click()

'This tells the user good luck with their homework and quits

MsgBox ("Good luck with your " & Homework & " hours of homework!")
End
End Sub

Private Sub cmdSudoku_Click()

'This takes the user to the sudoku application

frmMath.Hide
frmSudoku.Show
End Sub

