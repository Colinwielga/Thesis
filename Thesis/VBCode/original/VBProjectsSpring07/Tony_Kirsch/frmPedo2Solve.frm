VERSION 5.00
Begin VB.Form frmPedo2Solve 
   BackColor       =   &H00000000&
   Caption         =   "Case 4 Soultion"
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
      Left            =   4920
      ScaleHeight     =   1335
      ScaleWidth      =   3135
      TabIndex        =   10
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton cmddisplay2 
      BackColor       =   &H0080FF80&
      Caption         =   "Click to display pedophile typology you picked"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
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
      Left            =   9240
      ScaleHeight     =   1815
      ScaleWidth      =   2760
      TabIndex        =   8
      Top             =   3960
      Width           =   2760
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
   End
   Begin VB.PictureBox picwarning 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
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
      Height          =   2055
      Left            =   3840
      ScaleHeight     =   2055
      ScaleWidth      =   7215
      TabIndex        =   6
      Top             =   8400
      Width           =   7215
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmddisplay 
      BackColor       =   &H008080FF&
      Caption         =   "Click to display answer below"
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
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
      Height          =   855
      Left            =   2760
      ScaleHeight     =   855
      ScaleWidth      =   1935
      TabIndex        =   2
      Top             =   4080
      Width           =   1935
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
      Height          =   2055
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   2655
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
      Height          =   2175
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Label lbltitle 
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
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmPedo2Solve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form is if you complete the second pedophile case or case #4
'It has a few buttons on it that will help the user navigate: guess, display user
'profile, display correct profile, one to go back to case files, one to go back to the title screen



Private Sub cmdagain_Click()
'Takes the user back to the case files form so they can do another case
    frmPedo2Solve.Hide
    frmCasefiles.Show
End Sub

Private Sub cmdAnswer_Click()
'If you click on this button it will display what the correct answer is.
piccorrect.Cls 'clears the box of old junk
    piccorrect.Print "The correct profile is..." 'displays this
    piccorrect.Print Tab(5); "Preferential" 'displays this
    piccorrect.Print Tab(3); "Mysoped Molester" 'displays this
End Sub

Private Sub cmddisplay_Click()
picResult.Cls
    If check(14) / 2 = Int(check(14) / 2) Then
        picResult.Print "Situational"
    Else
        picResult.Print "Preferential"
    End If
End Sub

Private Sub cmddisplay2_Click()
picResult2.Cls
    If pedo2answer = "one" Then
        picResult2.Print "Regressed Molester"
    End If
    
    If pedo2answer = "two" Then
        picResult2.Print "Morally Indiscriminate Molester"
    End If
    
    If pedo2answer = "three" Then
        picResult2.Print "Sexually Indiscriminate Molester"
    End If
    
    If pedo2answer = "four" Then
        picResult2.Print "Inadequate Molester"
    End If
    
    If pedo2answer = "five" Then
        picResult2.Print "Mysoped Molester"
    End If
    
    If pedo2answer = "six" Then
        picResult2.Print "Fixated Molester"
    End If
    
End Sub

Private Sub cmdexit_Click()
'takes the user back to the title screen so that they can begin again with a new
'name or quit the program
    frmPedo2Solve.Hide
    frmTitleScreen.Show
End Sub

Private Sub cmdguess_Click()
'This is the guessing button. It is for anyone who thinks they know it all
    Dim Answer As String 'Declare my variable i will be using for this button
    'And what it equals which in this case will be whatever the user inputs
    Answer = InputBox("Please enter your guess. Remember to be exact in spelling. Remember it is case sensitive. Example: preferential inadequate molester", "Input please")
    'now if they type in exactly what i have listed they get a way to go.
    'if not they just get a sorry, and can click on the see right answer
    If Answer = "perferential mysoped molester" Then
        MsgBox "Well done, you should join my class", , "Congradulations" 'display if guessed right
    Else
        MsgBox "Sorry, but you are incorrect. Good guess though", , "Sorry you are wrong" 'display if guessed wrong
    End If
End Sub

Private Sub Form_Activate()
'Another form activated function
picwarning.Cls 'clears my picture box of any old stuff
    picwarning.Print "Always watch your children and know who they hang out with." 'Display this
    picwarning.Print "Make sure you know who approaches them and teach them" 'Display this
    picwarning.Print "how to not become a victim. This means You as well..." 'Display this
    picwarning.Print , nam ' Displays name of which user typed in earlier. Adds a personal touch i do believe.
End Sub

