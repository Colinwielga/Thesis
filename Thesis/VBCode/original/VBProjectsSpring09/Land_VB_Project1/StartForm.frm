VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00000000&
   Caption         =   "Start"
   ClientHeight    =   8790
   ClientLeft      =   2025
   ClientTop       =   1410
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   Picture         =   "StartForm.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   11340
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuiz 
      BackColor       =   &H00000080&
      Caption         =   "Click to quiz your knowledge on the Twilight series."
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton cmdMatch 
      BackColor       =   &H00000080&
      Caption         =   "Click to play a matching game."
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton cmdCharacters 
      BackColor       =   &H00000080&
      Caption         =   "Click to learn more about the characters in Twilight."
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdBooks 
      BackColor       =   &H00000080&
      Caption         =   "Click to Review the four books in the Twilight series."
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   6
      Top             =   0
      Width           =   0
   End
   Begin VB.Label lblTwilight 
      BackColor       =   &H000000C0&
      Caption         =   "Welcome to the Twilight"
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmStart
'Author: Mollie Land
'Date Written: 3/15/09
'Objective: This is the initial start form where the user chooses which activity
'they would like to do in the program
'Purpose of Project: The purpose of this project is to allow the user to learn more
'about the Twilight series.  With this project once can learn about the 4 books as well
'as the main characters in the books.  The user can also quiz himself/herself on
'the books and characters to test their knowledge on the series

Private Sub cmdBooks_Click()
    'show book form, hide the start form
    frmBooks.Show
    frmStart.Hide
End Sub

Private Sub cmdCharacters_Click()
    'show characters form, hide the start form
    frmCharacters.Show
    frmStart.Hide
End Sub

Private Sub cmdMatch_Click()
    'show match form, hide the start form
    frmMatchPictures.Show
    frmStart.Hide
End Sub

Private Sub cmdQuit_Click()
    'quit the program
    End
End Sub

Private Sub cmdQuiz_Click()
    'show quiz form, hide the start form
    frmQuiz.Show
    frmStart.Hide
End Sub


