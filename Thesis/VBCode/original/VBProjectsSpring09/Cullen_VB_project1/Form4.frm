VERSION 5.00
Begin VB.Form frmquiz 
   BackColor       =   &H000000C0&
   Caption         =   "Form4"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form4"
   ScaleHeight     =   10440
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   3600
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   5835
      TabIndex        =   22
      Top             =   2040
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bonus: Please select one "
      Height          =   855
      Left            =   4200
      TabIndex        =   21
      Top             =   8280
      Width           =   4455
      Begin VB.CheckBox Check2 
         Caption         =   "Billy"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Benny"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "quit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      TabIndex        =   19
      Top             =   9360
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "return"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   18
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Compute Score!!"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      TabIndex        =   17
      Top             =   9360
      Width           =   2175
   End
   Begin VB.TextBox txtbox5 
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox txtbox4 
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtbox3 
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtbox2 
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtbox1 
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
      Caption         =   "What is the name of the Bulls Mascot?"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000C0&
      Caption         =   "Bonus: What is the name of the Bulls Mascot?"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   7080
      Width           =   4815
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "5. The Stud"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9840
      TabIndex        =   16
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "4. Stevie Wonder"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9840
      TabIndex        =   15
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "3. The Worm"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9840
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "2. Air"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9840
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      Caption         =   $"Form4.frx":2812D
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   8775
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "1. The Big Aussie"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   9840
      TabIndex        =   11
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Dennis Rodman"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Luc Longley"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H000000C0&
      Caption         =   "Chicago Bulls Nickname Quiz"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   1800
      TabIndex        =   8
      Top             =   240
      Width           =   9255
   End
   Begin VB.Label lblsteve 
      BackColor       =   &H80000012&
      Caption         =   "Steve Kerr"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblmj 
      BackColor       =   &H80000012&
      Caption         =   "Michael Jordan"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lbljud 
      BackColor       =   &H80000012&
      Caption         =   "Jud Buechler"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmquiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    'Chicago Bulls (Chicagobulls.vbp)
    'frmquiz(frmquiz.frm)
    'Written by: Brian Cullen
    'Written on: March 16, 2008
    'Objective: This form allows the user to take a quiz on nicknames of the chicago bulls by
    'viewing a photo of each player.

Private Sub cmdcompute_Click()
Dim answer1 As Integer, answer2 As Integer, answer3 As Integer, answer4 As Integer, answer5 As Integer
Dim sum As Single
Dim sum2 As Single
answer1 = txtbox1.Text
answer2 = txtbox2.Text
answer3 = txtbox3.Text
answer4 = txtbox4.Text
answer5 = txtbox5.Text
sum = 0

Dim score As Single
If answer1 = "5" Then
sum = sum + 1
End If
If answer2 = "2" Then
sum = sum + 1
End If
If answer3 = "4" Then
sum = sum + 1
End If
If answer4 = "1" Then
sum = sum + 1
End If
If answer5 = "3" Then
sum = sum + 1
End If

score = sum / 5
sum2 = sum * 20

Select Case sum2
Case Is >= 100
    MsgBox ("Your score is " & FormatPercent(score) & " you are amazing")
Case 80 To 99
    MsgBox ("Your score is " & FormatPercent(score) & " you must be a bulls fan")
Case 40 To 79
    MsgBox ("Your score is " & FormatPercent(score) & " very weak performance")
Case 0 To 39
    MsgBox ("Your score is" & FormatPercent(score) & " you must go to Saint Thomas")
End Select


Dim Correctanswer As String
If Check1.Value = 1 Then
MsgBox ("You got the bonus question Correct!")
ElseIf Check2.Value = 0 Then
MsgBox ("Sorry, the Chicago Bulls mascot's name is benny")
End If





End Sub


Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
frmquiz.Hide
frmmainpage.Show

End Sub



