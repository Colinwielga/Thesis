VERSION 5.00
Begin VB.Form frmTrivia 
   BackColor       =   &H00FF0000&
   Caption         =   "Trivia"
   ClientHeight    =   7035
   ClientLeft      =   2520
   ClientTop       =   1920
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10410
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FF00&
      Height          =   1095
      Left            =   2880
      ScaleHeight     =   1035
      ScaleWidth      =   7035
      TabIndex        =   11
      Top             =   5400
      Width           =   7095
   End
   Begin VB.CommandButton cmdcompute 
      BackColor       =   &H0000FF00&
      Caption         =   "Click Here to Find your Mickey Mouse Club ID"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H0000FF00&
      Height          =   975
      Left            =   3000
      TabIndex        =   8
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H000080FF&
      Caption         =   "You Ready? Click here to   see how sharp your Disney Trivia Skills are!"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   3135
   End
   Begin VB.PictureBox cmd 
      BackColor       =   &H008080FF&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdTickets 
         BackColor       =   &H0000C000&
         Caption         =   "Buy Your Tickets Now!"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CommandButton cmdTop 
         BackColor       =   &H000080FF&
         Caption         =   "Top 10 Disney Animated Movies Of All Time"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00800080&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CommandButton cmdGiftShop 
         BackColor       =   &H00FF0000&
         Caption         =   "Gift Shop "
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdIntro 
         BackColor       =   &H000000FF&
         Caption         =   "Main Page"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   6840
      Picture         =   "DisneyLand Trivia.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Please Enter your First name, Middle Intitial, and Last name.  Example (Kelly M Holmseth)"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is to allow the user to play the Disney game by clicking on a command button


Private Sub cmdBegin_Click()
Dim TriviaArray(1 To 15) As String  'loads the questions and answers at the same time.
Dim AnswerArray(1 To 15) As String
Dim x As String
Dim Sum As Single
Dim Count As Integer
Dim I As Single
Dim y As String
Sum = 0
Count = 0
Open App.Path & "\Questions.txt" For Input As #1        'Open file
    Do Until EOF(1)
        Sum = Sum + 1  'adds one to the sum or counter and just keeps the questions rolling
        Input #1, TriviaArray(Sum), AnswerArray(Sum)  'imputs both questions and ansers simultaneously, we made an array with both questions in a textfile called questions.txt making it very easy to go and change the questions to fit your audience.
    Loop
Close #1
Sum = 1
MsgBox "Please Use Lower-Case One Word Answers", , "Guidelines"  'these 2 message boxes pop up before the questions start in order to explain directions to the user.
MsgBox "This quiz is all about Disney characters, movies, and numbers.  We will give you a description, and you tell us the answer.  Have fun!"
    For Sum = 1 To 15
    y = TriviaArray(Sum)
    x = InputBox(y, "Question")
    If x = AnswerArray(Sum) Then 'if answer input by the user matches the answer stored in the array then 1 will be added to the final count of questions answered correctly.
        Count = Count + 1 'keeps total # correct tally.
    End If
Next Sum

Select Case Count  'different cases of numbers correct print out different comments.  Notice the spacing in cases below.
Case Is >= 15
    MsgBox "Your Score is: " & Count & "/15,  Perfect!", , "Score"
Case 11 To 14
    MsgBox "Your Score is: " & Count & "/15,  Excellent!", , "Score"
Case 7 To 10
    MsgBox "Your Score is: " & Count & "/15,  Good Job!", , "Score"
Case 4 To 6
    MsgBox "Your Score is: " & Count & "/15,  Atleast you Tried!", , "Score"
Case 1 To 3
    MsgBox "Your Score is: " & Count & "/15,  You better go watch a few more Disney Movies!", , "Score"
Case Else
    MsgBox "Your Score is: " & Count & "/15,  Have you ever even heard of a Disney Movie?!", , "Score"
End Select



'MsgBox "Your score is:" & Count, , "Excellent!"
frmGiftShop.Hide            'Allows the User to go to the Trivia form
frmIntro.Hide
frmTrivia.Show
End Sub

Private Sub cmdcompute_Click()
Dim wholeName As String, ID As String
Dim N As Integer
Dim first As String, middle As String, last As String
picResults.Cls
wholeName = txtName.Text  'this whole process takes Danny M Hansen and turns it into DMHansen, kind of like our user name for SJU and CSB
N = InStr(wholeName, " ")
first = Left(wholeName, N - 1)  'notice left and the number of spaces needing to move over
last = Right(wholeName, Len(wholeName) - (N + 2))
middle = Mid(wholeName, N + 1, 1)
ID = Left(first, 1) & middle & Left(last, 6) 'takes 6 letters from your last name

picResults.Print "Your Mickey Mouse Club User Name is", ID 'nothing really happens here it is just pretend.
picResults.Print
picResults.Print "Be sure to use this name when you email us @ www.mickeymouseclub.com"
picResults.Print "to be entered into our prize giveaways. Have a great day "; ID; "!!!!"

End Sub

Private Sub cmdGiftShop_Click()
frmTrivia.Hide      'Allow the user to go to the GiftShop form
frmIntro.Hide
frmGiftShop.Show
frmTop.Hide
frmTickets.Hide

End Sub

Private Sub cmdIntro_Click()
frmTrivia.Hide      'Allows the user go to the Intro form
frmGiftShop.Hide
frmIntro.Show
frmTop.Hide
frmTickets.Hide

End Sub

Private Sub cmdQuit_Click()     'Allows the user to quit the program
End
End Sub

Private Sub cmdTickets_Click()
frmTrivia.Hide      'Allows the user to go to the tickets form
frmIntro.Hide
frmGiftShop.Hide
frmTop.Hide
frmTickets.Show

End Sub

Private Sub cmdTop_Click()
frmTrivia.Hide
frmTop.Show
frmGiftShop.Hide
frmTickets.Hide
frmIntro.Hide

End Sub

