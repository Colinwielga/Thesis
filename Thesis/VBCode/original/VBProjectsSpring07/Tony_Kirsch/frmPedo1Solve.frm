VERSION 5.00
Begin VB.Form frmPedo1Solve 
   BackColor       =   &H00000000&
   Caption         =   "Case 3 Soultion"
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
      Left            =   5160
      ScaleHeight     =   1335
      ScaleWidth      =   2655
      TabIndex        =   10
      Top             =   4080
      Width           =   2655
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
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
      Left            =   10080
      ScaleHeight     =   1815
      ScaleWidth      =   2760
      TabIndex        =   8
      Top             =   4320
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
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
      Left            =   4560
      ScaleHeight     =   2055
      ScaleWidth      =   6735
      TabIndex        =   6
      Top             =   8640
      Width           =   6735
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   2295
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
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
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
      Height          =   975
      Left            =   2760
      ScaleHeight     =   975
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
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
      Left            =   11640
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
      Left            =   1680
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
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmPedo1Solve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form is if you complete the first pedophile case or case #3
'It has a few cool buttons on it that will help you navigate, one to guess, one
'to display your profile, and one to display to right profile.



Private Sub cmdagain_Click()
'takes the user back to the case files form so they can do another one
    frmPedo1Solve.Hide
    frmCasefiles.Show
    
End Sub

Private Sub cmdAnswer_Click()
'If you click on the correct answer button you get the right profile
piccorrect.Cls 'Clears out the picture box of old data
    piccorrect.Print "The correct profile is..." 'Displays this message
    piccorrect.Print Tab(5); "Situational" 'Displays this message
    piccorrect.Print Tab(3); "Regressed Molester" 'Displays this message
End Sub

Private Sub cmddisplay_Click()
picResult.Cls
    If check(13) / 2 = Int(check(13) / 2) Then
        picResult.Print "Situational"
    Else
        picResult.Print "Preferential"
    End If
    
End Sub


Private Sub cmddisplay2_Click()
picResult2.Cls
    If pedo1answer = "one" Then
        picResult2.Print "Regressed Molester"
    End If
    
    If pedo1answer = "two" Then
        picResult2.Print "Morally Indiscriminate Molester"
    End If
    
    If pedo1answer = "three" Then
        picResult2.Print "Sexually Indiscriminate Molester"
    End If
    
    If pedo1answer = "four" Then
        picResult2.Print "Inadequate Molester"
    End If
    
    If pedo1answer = "five" Then
        picResult2.Print "Mysoped Molester"
    End If
    
    If pedo1answer = "six" Then
        picResult2.Print "Fixated Molester"
    End If
End Sub

Private Sub cmdexit_Click()
'This takes you to the title screen which allows the user to exit
    frmPedo1Solve.Hide
    frmTitleScreen.Show
End Sub

Private Sub cmdguess_Click()
'This button i made for a guessing game factor. If a person thinks they know what it
'the answer is then they can click and type it in. however the spelling and case
'must be the exact same
Dim Answer As String 'Declared the variable
    'An input box pops up and they type in what they think the answer is
    Answer = InputBox("Please enter your guess. Remember to be exact in spelling. Remember it is case sensitive. Example: preferential inadequate molester", "Input please")
    'if what they type in is what what i wrote exactly then they get a congradulations.
    'otherwise they simply get a negative respsone and are free to just click on the answer
    If Answer = "situational regressed molester" Then
        MsgBox "Well done, you should join my class", , "Congradulations" 'If they do well
    Else
        MsgBox "Sorry, but you are incorrect. Good guess though", , "Sorry you are wrong" 'if they guess wrong
    End If
End Sub

Private Sub Form_Activate()
'I made this as kind of a little warning because i have how to avoide rapists on the
'other forms i included this here. it pops up when the form is activated
picwarning.Cls 'clears the box of old junk
    picwarning.Print "Always watch your children and know who they hang out with." 'displays this
    picwarning.Print "Make sure you know who approaches them and teach them" 'displays this
    picwarning.Print "how to not become a victim. This means You as well..." 'displays this
    picwarning.Print , nam 'Displays the name that the user inputed at the beginning of the program.
                            'make it more personal.
End Sub

