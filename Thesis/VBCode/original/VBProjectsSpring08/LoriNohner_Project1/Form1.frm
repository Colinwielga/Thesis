VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   Caption         =   "Welcome!"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   FillColor       =   &H80000000&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdWelcome 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Welcome to Disney Castle!! "
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   6735
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Project
'Intro
'Lori Nohner
'Written- March 17, 2008
'Objective of project-  This is just for fun.  The user is able to look at pictures of
    'the disney Princesses and learn about Disney villains.  They can play a game, buy
    'souviners and look at a list of Disney movies.

Option Explicit


Private Sub cmdExit_Click()
    End 'Stops program
End Sub

Private Sub cmdWelcome_Click()
    UserName = InputBox("Enter Your Name.", "Welcome!") 'asks user for a name and stores it
    frmIntro.Hide 'hides Intro page
    frmDisneyCastle.Show 'makes next form visible
    MsgBox "Welcome to the Disney Castle, " & UserName & ".  Feel free to explore!", , "Look Around."
    
End Sub


