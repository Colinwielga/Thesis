VERSION 5.00
Begin VB.Form frmPiano2 
   Caption         =   "Learning the Piano 2"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmPiano2.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinished 
      BackColor       =   &H0080FF80&
      Caption         =   "Finished"
      Height          =   1455
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back To Main Page"
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox txt10 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   9960
      TabIndex        =   32
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txt9 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Height          =   495
      Left            =   9960
      TabIndex        =   31
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txt8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Height          =   495
      Left            =   9960
      TabIndex        =   30
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txt7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   9960
      TabIndex        =   29
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txt6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   9960
      TabIndex        =   28
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txt5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   8040
      TabIndex        =   22
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txt4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   8040
      TabIndex        =   19
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   8040
      TabIndex        =   18
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox picKeyboard1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   2160
      Picture         =   "frmPiano2.frx":2DBBC2
      ScaleHeight     =   2430
      ScaleWidth      =   4920
      TabIndex        =   2
      Top             =   4440
      Width           =   4980
      Begin VB.Label lbl10 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "10."
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "9."
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lbl8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "8."
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lbl7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "7."
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "6."
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lbl5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "5."
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lbl4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "4."
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbl3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "3."
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "2."
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "1."
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1680
         Width           =   375
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "10."
      Height          =   375
      Left            =   9360
      TabIndex        =   27
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "9."
      Height          =   375
      Left            =   9360
      TabIndex        =   26
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "8."
      Height          =   375
      Left            =   9360
      TabIndex        =   25
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "7."
      Height          =   375
      Left            =   9360
      TabIndex        =   24
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "6."
      Height          =   375
      Left            =   9360
      TabIndex        =   23
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "5."
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "4."
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "3."
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "2."
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "1."
      Height          =   375
      Left            =   7560
      TabIndex        =   13
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   $"frmPiano2.frx":302AB4
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   1560
      TabIndex        =   1
      Top             =   2280
      Width           =   9495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Notes on the Piano"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   11415
   End
End
Attribute VB_Name = "frmPiano2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This page quizes the user about the keys of the piano asking for the letter names of the piano keys, and asking only 10 questions
'The user inputs there answers into textboxes and if answers are correct then 1 point is added to PianoPoints
'Once finished their score is given to them via message box

Private Sub cmdBack_Click()     'This button changes forms to frmLessonMainPage
    frmPiano2.Hide                  'this hides frmPiano2
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

'This button checks all of the input boxes to see if the user entered the correct answer only in lower case letters, upper case would not be accepted
'there are If statement to check if the answers are correct, if they are 1 point is added to PianoPoints Public variable, if not nothing happens
Private Sub cmdFinished_Click()
PianoPoints = 0                     'sets value of PianoPoints = 0
    If txt1.Text = "f" Then                         'If statement checks the correct answer for txt1.Text
        PianoPoints = PianoPoints + 1           'if correct adds 1 to value of PianoPoints
    End If
    If txt2.Text = "b" Then                         'If statement checks the correct answer for txt2.Text
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    If txt3.Text = "d#" Or txt3.Text = "eb" Then    'If statement checks the correct answer for txt3.Text, in this question there are 2 correct answers so it looks for both
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    If txt4.Text = "ab" Or txt4.Text = "g#" Then    'If statement checks the correct answer for txt4.Text, in this question there are 2 correct answers so it looks for both
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    If txt5.Text = "d" Then                         'If statement checks the correct answer for txt5.Text
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    If txt6.Text = "c" Then                         'If statement checks the correct answer for txt6.Text
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    If txt7.Text = "a#" Or txt7.Text = "bb" Then    'If statement checks the correct answer for txt7.Text, in this question there are 2 correct answers so it looks for both
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    If txt8.Text = "e" Then                         'If statement checks the correct answer for txt8.Text
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    If txt9.Text = "g" Then                         'If statement checks the correct answer for txt9.Text
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    If txt10.Text = "a" Then                        'If statement checks the correct answer for txt10.Text
        PianoPoints = PianoPoints + 1               'if correct adds 1 to value of PianoPoints
    End If
    MsgBox "You got " & PianoPoints & " points!!! Congratulations " & NameGiven & "!!!", , "Your Score" 'Displays a message box with the person's name and score from this quiz
End Sub
