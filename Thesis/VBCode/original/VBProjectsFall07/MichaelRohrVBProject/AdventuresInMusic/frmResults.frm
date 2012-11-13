VERSION 5.00
Begin VB.Form frmResults 
   BackColor       =   &H00400000&
   Caption         =   "Results"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      Height          =   4095
      Left            =   1200
      ScaleHeight     =   4035
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdClickMe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to See an Overall Score from the different lessons"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Results from the Program Quizes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This page is used as a way for the user to see what they got in all of the quizes that they took in this program if they so choose,
'if the user did not take a quiz then a line shows up, Have not taken...quiz.
'it then gives the option of quiting the program from there instead of going back to the main page to do so

Private Sub cmdBack_Click()     'This button changes forms to frmLessonMainPage
    frmResults.Hide                 'this hides frmResults
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

'This button displays the scores of the quizes taken by the user,
'if no quiz was taken then it displays a message saying quiz not taken
'it finds out if the quiz was taken by asking if the score is > 0,
'I would be assuming that every student taking this quiz would at least get 1 point, how great on my part!
Private Sub cmdClickMe_Click()
    picResults.Print "Here Are the Results of Your Lessons" & NameGiven & ":"
    picResults.Print "____________________________________"
    If PointsMusic > 0 Then
        picResults.Print "Music Basics Quiz: You Got " & PointsMusic & " for your score."
    Else
        picResults.Print "Have Not Taken Music Basics Quiz"
    End If
    If PianoPoints > 0 Then
        picResults.Print "Piano Skills Quiz: You Got " & PianoPoints & " for your score."
    Else
        picResults.Print "Have Not Taken Piano Skills Quiz"
    End If
    If TreblePoints > 0 Then
        picResults.Print "Treble Clef Quiz: You got " & TreblePoints & " for your score."
    Else
        picResults.Print "Have Not Taken Treble Clef Quiz"
    End If
    If BassPoints > 0 Then
        picResults.Print "Bass Clef Quiz: You got " & BassPoints & " for your score."
    Else
        picResults.Print "Have Not Taken Bass Clef Quiz"
    End If
End Sub

'This button ends the program and displays a farewell greeting before closing
Private Sub cmdQuit_Click()
    MsgBox "Goodbye " & NameGiven & "Have a Great Day!", , "Have a Great Day!"
End
End Sub
