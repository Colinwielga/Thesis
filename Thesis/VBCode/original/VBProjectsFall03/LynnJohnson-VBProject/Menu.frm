VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00808000&
   Caption         =   "Menu (Lynn Johnson)"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form3"
   ScaleHeight     =   8295
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   4020
      TabIndex        =   2
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdplaygame 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Start a Game of Memory!"
      Default         =   -1  'True
      Height          =   1095
      Left            =   2760
      MaskColor       =   &H80000010&
      TabIndex        =   0
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Welcome to the Game of Memory!"
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   6135
   End
   Begin VB.Label lblmenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Main Menu"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : TheMemoryGame (Lynn Johnson - VB Project)
'Form Name : frmMenu (Menu.frm)
'Author: Lynn Johnson
'Date Written: October 29, 2003
'Purpose of Project: The overall purpose is to allow the user to have fun by playing the
                'mind-challenging game of memory.  This game allows the user to
                'play five games in which the cards are shuffled in different ways.
                'Scores are then formulated based on the amount of mismatched
                'cards found before the end of the game.
                
'Purpose of Form: The purpose of this form is to serve as the home-base
                'for the project.  This is the Main Menu.  It takes the
                'user from this form to a game form.

'Option Explicit allows the user to declare
'varibles that will be used throughout the whole form.
Option Explicit

Private Sub cmdplaygame_Click()
    'Move from Menu form to game form
    Dim game As Integer
        game = InputBox("Enter a game number between the numbers 1 and 5")
        
    Select Case game
        Case 1
            frmGame1.Show
            frmMenu.Hide
        Case 2
            frmGame2.Show
            frmMenu.Hide
        Case 3
            frmGame3.Show
            frmMenu.Hide
        Case 4
            FrmGame4.Show
            frmMenu.Hide
        Case 5
            FrmGame5.Show
            frmMenu.Hide
        Case Else
            MsgBox "That number is not between 1 and 5.  Pick another number", , "Error"
    End Select
    
End Sub

Private Sub cmdquit_Click()
    End
    
End Sub

Private Sub Form_Load()

    strPath = "N:\CS130\handin\Lynn Johnson - VB Project\AnimalCards.txt"

End Sub

