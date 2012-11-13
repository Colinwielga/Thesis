VERSION 5.00
Begin VB.Form frmNascar 
   BackColor       =   &H00800000&
   Caption         =   "Nascar"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000D&
      Caption         =   "Exit"
      Height          =   615
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdWelcome 
      Caption         =   "Welcome to the Nascar experience of a life time!   Click here to enter."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   0
      Top             =   3000
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   1200
      Picture         =   "frmNascar.frx":0000
      Top             =   360
      Width           =   9000
   End
End
Attribute VB_Name = "frmNascar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Introduction to NASCAR
'Form Nascar
'Colin Roberts and Luke Hommerding
'Written 10/18/09
'The purpose of this visual basic project is to apply all the concepts learned in class
'into a visual basic program that shares the knowledge of what interests the designers.
'We chose NASCAR because Colin and Luke are both big fans of the sport and thought that
'using visual basic would be an interesting way to teach people about the sport.
'People who use this program will learn much information about the sport of NASCAR and
'how visual basic works.
Option Explicit
'asks user for their name
Private Sub cmdWelcome_Click()
'enter users name
    UserName = InputBox("What is your name?", "You are now a part of Nascar!")
    'takes user to the main menu
    frmMain.Show
    frmNascar.Hide
    'greetings
    MsgBox "Welcome to the Nascar experience, " & UserName & "!"

End Sub
'quits the program
Private Sub cmdExit_Click()
End
End Sub
