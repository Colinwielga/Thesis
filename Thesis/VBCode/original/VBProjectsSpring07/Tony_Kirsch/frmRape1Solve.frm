VERSION 5.00
Begin VB.Form frmRape1Solve 
   BackColor       =   &H00000000&
   Caption         =   "Case One Solution"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picResult2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   5760
      ScaleHeight     =   1335
      ScaleWidth      =   3975
      TabIndex        =   13
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton cmddisplay2 
      BackColor       =   &H0080FF80&
      Caption         =   "Click to display the rapist typology you picked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox picarrest 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6240
      Picture         =   "frmRape1Solve.frx":0000
      ScaleHeight     =   2415
      ScaleWidth      =   2655
      TabIndex        =   11
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton cmdAnswer 
      BackColor       =   &H00FF8080&
      Caption         =   "Click to display correct answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox piccorrect 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   10680
      ScaleHeight     =   1815
      ScaleWidth      =   2760
      TabIndex        =   9
      Top             =   4320
      Width           =   2760
   End
   Begin VB.CommandButton cmdagain 
      BackColor       =   &H00FFFF00&
      Caption         =   "Return to case files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8520
      Width           =   2535
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C000C0&
      Caption         =   "Return to Title Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8520
      Width           =   2655
   End
   Begin VB.CommandButton cmdhere2 
      BackColor       =   &H000000FF&
      Caption         =   "Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9480
      Width           =   735
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3120
      ScaleHeight     =   1335
      ScaleWidth      =   2055
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmddisplay 
      BackColor       =   &H008080FF&
      Caption         =   "Click to display the profile you created"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton cmdguess 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Want to take a guess?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lbldes2 
      BackColor       =   &H00000000&
      Caption         =   "to learn about methods of avoidance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   9480
      Width           =   4215
   End
   Begin VB.Label lblclick2 
      BackColor       =   &H00000000&
      Caption         =   "Click"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   9480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Case Solution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmRape1Solve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form is if you complete the first rape case or case #1
'It has a few buttons on it that will help the user navigate: guess, display user
'profile, display correct profile, one to go back to case files, one to go back to the title screen

Private Sub cmdagain_Click()
'Takes the user back to the case files to do another one
    frmRape1Solve.Hide
    frmCasefiles.Show
    
End Sub

Private Sub cmdAnswer_Click()
'This button displays what the correct profile is
piccorrect.Cls 'Gets rid of anything that might be in the picture box
    piccorrect.Print "The correct profile is..." 'displays this
    piccorrect.Print Tab(5); "Disorganized" 'displays this
    piccorrect.Print Tab(3); "Anger Retaliatory" 'displays this
End Sub

Private Sub cmddisplay_Click()
'declare all variables for the button. There are a few.
    Dim pos As Integer, sumone As Integer, sumtwo As Integer, sumthree As Integer
    Dim sumfour As Integer, sumfive As Integer, sumsix As Integer
    picResult.Cls 'clears out picture box in case you press it twice.
    For pos = 1 To 3 'i declare this to be pos because the first three boxes are what define disorganized
        If check(pos) / 2 = Int(check(pos) / 2) Then 'quantative formula using even numers to determine the value of the check box
            sumone = sumone 'made up varibale to keep it straight.
        Else
            sumone = sumone + 1 'to keep using the an even odd strategy
        End If
    Next pos 'loop until all three have been checked
    If sumone >= 3 Then 'if the person selected all three then
        picResult.Print "Disorganized" 'they get to display the right answer
    End If
    
    
    For pos = 4 To 6 'i declare this to be pos because the last three boxes are what define organized
        If check(pos) / 2 = Int(check(pos) / 2) Then 'quantative formula using even numers to determine the value of the check box
            sumtwo = sumtwo 'made up varibale to keep it straight, but changed to not confuse.
        Else
            sumtwo = sumtwo + 1 'to keep using the an even odd strategy
        End If
    Next pos 'loop until all three have been checked
    If sumtwo >= 3 Then 'if a person checked off all three of the right side boxes then
        picResult.Print "Organized" 'they get to display the right answer
    End If
    
    If sumtwo < 3 And sumone < 3 Or sumtwo = 3 And sumone > 1 Or sumone = 3 And sumtwo > 1 Then 'long conditional essentially saying if it is not exact the response will be inconclusive.
        picResult.Cls 'clears out the picture box
        picResult.Print "Inconclusive" 'print out should the users information meet the criteria above.
    End If
    
End Sub

Private Sub cmddisplay2_Click()
picResult2.Cls 'clears picture box
    If rape1answer = "one" Then 'using the values assigned it prints the corresponding answer.
        picResult2.Print "Power Assertive"
    End If
    
    If rape1answer = "two" Then 'using the values assigned it prints the corresponding answer.
        picResult2.Print "Power Reassurance"
    End If
    
    If rape1answer = "three" Then 'using the values assigned it prints the corresponding answer.
        picResult2.Print "Anger retaliatory"
    End If
    
    If rape1answer = "four" Then 'using the values assigned it prints the corresponding answer.
        picResult2.Print "Anger Excitatory"
    End If
End Sub

Private Sub cmdexit_Click()
'Takes the user back to the title screen to do start over or quit
    frmRape1Solve.Hide
    frmTitleScreen.Show
End Sub

Private Sub cmdguess_Click()
'This button is for guessing the right answer
    Dim Answer As String 'Declare my variable
    'the data is recieved through an input box
    Answer = InputBox("Please enter your guess. Remember to be exact in spelling. Remember answer is case sensitive. Example: Organized Power Assertive", "Input please")
    'If they type in the answer just right then they win, if not no harm no foul
    If Answer = "disorganized anger retaliatory" Then
        MsgBox "Well done, you should join my class", , "Congratulations" 'display if right
    Else
        MsgBox "Sorry, but you are incorrect. Good guess though", , "Sorry you are wrong" 'Display if wrong
    End If
    
End Sub

Private Sub cmdhere2_Click()
'Takes the user to a page with advice on how to get away from rapists
    frmRape1Solve.Hide
    frmavoid.Show
End Sub


Private Sub Form_Activate()
'Clears the picture box as soon as the form is activated
 picResult.Cls
End Sub
