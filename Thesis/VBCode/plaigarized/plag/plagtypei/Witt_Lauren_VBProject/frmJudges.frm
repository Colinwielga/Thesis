VERSION 5.00
Begin VB.Form frmJudges
   BackColor       =   &H00000000&
   Caption         =   "Judges Panel"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4
      Height          =   1695
      Left            =   1080
      Picture         =   "frmJudges.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.PictureBox Picture3
      Height          =   1575
      Left            =   1080
      Picture         =   "frmJudges.frx":BCFA
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   4320
      Width           =   1935
   End
   Begin VB.PictureBox Picture2
      Height          =   1575
      Left            =   1080
      Picture         =   "frmJudges.frx":17014
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdGo
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click here for their opinions"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.PictureBox picResults
      BackColor       =   &H00FFC0FF&
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   4560
      ScaleHeight     =   4035
      ScaleWidth      =   8595
      TabIndex        =   1
      Top             =   2160
      Width           =   8655
   End
   Begin VB.Label lblj3
      BackColor       =   &H00000000&
      Caption         =   "Judge 3 "
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3360
      TabIndex        =   8
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label1
      BackColor       =   &H00000000&
      Caption         =   "Judge 2"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3360
      TabIndex        =   7
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblj1
      BackColor       =   &H00000000&
      Caption         =   "Judge 1"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3240
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lbljudges
      BackColor       =   &H00FF80FF&
      Caption         =   "                Judges panel"
      BeginProperty Font
         Name            =   "Bookman Old Style"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   13455
   End
End
Attribute VB_Name = "frmJudges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdGo_Click()    ' comments here and there and stuff
Dim YourScore As Integer    ' comments here and there and stuff


YourScore = InputBox("Enter your team's score", "Score")    ' comments here and there and stuff

Select Case YourScore    ' comments here and there and stuff


    Case Is >= 45    ' comments here and there and stuff

        picResults.Print "Judge Number 1 says: 'OMG! You are fantastic!'"    ' comments here and there and stuff
        picResults.Print "Judge Number 2 says: 'I'm speechless'"    ' comments here and there and stuff


        picResults.Print "Judge Number 3 says: 'Great work!'"    ' comments here and there and stuff
    Case 35 To 44    ' comments here and there and stuff
        picResults.Print "Judge Number 1 says: 'OMG! You are pretty good!'"    ' comments here and there and stuff
        picResults.Print "Judge Number 2 says: 'I'm inpressed'"    ' comments here and there and stuff

        picResults.Print "Judge Number 3 says: 'Good!'"    ' comments here and there and stuff



    Case 22 To 34    ' comments here and there and stuff
        picResults.Print "Judge Number 1 says: 'Not bad!'"    ' comments here and there and stuff
        picResults.Print "Judge Number 2 says: 'You did okay'"    ' comments here and there and stuff
        picResults.Print "Judge Number 3 says: 'I think you could have done better but still not bad!'"    ' comments here and there and stuff
    Case 15 To 24    ' comments here and there and stuff

        picResults.Print "Judge Number 1 says: 'Kinda bad!'"    ' comments here and there and stuff
        picResults.Print "Judge Number 2 says: 'Ya, you can improve'"    ' comments here and there and stuff
        picResults.Print "Judge Number 3 says: 'Not so hot!'"    ' comments here and there and stuff
    Case 5 To 14    ' comments here and there and stuff


        picResults.Print "Judge Number 1 says: 'Bad!'"    ' comments here and there and stuff
        picResults.Print "Judge Number 2 says: 'You really need to work on your relationship'"    ' comments here and there and stuff
        picResults.Print "Judge Number 3 says: 'Do you really live together?!'"    ' comments here and there and stuff
    Case Else    ' comments here and there and stuff
        picResults.Print "Judge Number 1 says: 'I think you know how you did!'"    ' comments here and there and stuff
        picResults.Print "Judge Number 2 says: 'I am speechless! You suck!'"    ' comments here and there and stuff



        picResults.Print "Judge Number 3 says: 'I know you better than your roommate does!'"    ' comments here and there and stuff


End Select    ' comments here and there and stuff

End Sub    ' comments here and there and stuff
