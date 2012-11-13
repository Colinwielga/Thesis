VERSION 5.00
Begin VB.Form Hangman 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   FillColor       =   &H000040C0&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   Picture         =   "Hangman.frx":0000
   ScaleHeight     =   10020
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back to Main Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9840
      TabIndex        =   2
      Top             =   5520
      Width           =   3375
   End
   Begin VB.CommandButton cmdhangman2 
      Caption         =   "Hangman #2"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9840
      TabIndex        =   1
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton cmdhangman1 
      Caption         =   "Hangman #1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9840
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "Hangman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'Form Hangman
'Joel Coleman
'March 29, 2008
'To give the user a chance to select which game to play first
'I used the .show and .hide functions
Private Sub cmdhangman1_Click()
Hangman1.Show
Hangman.Hide
End Sub

Private Sub cmdhangman2_Click()
Hangman.Hide
Hangman2.Show
End Sub

Private Sub Command1_Click()
Hangman.Hide
frmMainpage.Show
End Sub
