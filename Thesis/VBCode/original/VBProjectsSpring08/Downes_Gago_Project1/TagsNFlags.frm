VERSION 5.00
Begin VB.Form TagsNFlags 
   BackColor       =   &H00400000&
   Caption         =   "Tag The Flags"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   5520
      Width           =   2895
   End
   Begin VB.CommandButton cmdFlags 
      Cancel          =   -1  'True
      Caption         =   "LOAD FLAGS"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Image picImage 
      Height          =   6615
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "TagsNFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: TagsNFlags.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  This Form is a game for the user.
'He attempt to answer what flag is pictured on the form
'by typing the answer into an input box.  The Score is
'then calculated for the number of questions he answers

Option Explicit
    'Dim the variables for the whole form
'Hides the TagsNFlags Form and brings the user back to the South America Form
Private Sub cmdBack_Click()
TagsNFlags.Hide
SouthAmerica.Show
End Sub
'Flags pop up into an image box and a message box appears so that the user can type in the answers to the flags
Private Sub cmdFlags_Click(index As Integer)
Dim Flag As String, Score As Single, country(1 To 14) As String
Dim FinalScore As Single
'Dim variables for just this program
j = 0
ctr = 0     'the counters "J" and ctr are set to zero
            'A data file is opened with the names of the countries and is put into an array
    Open App.Path & "\Flags\Countries.txt" For Input As #1
    'Using a Do While Loop, the data is put into an array
Do While Not EOF(1)
    ctr = ctr + 1
        Input #1, country(ctr)
                'In this Do While Loop, a picture is loaded into the picImage box every time the user answers a question
            picImage.Picture = LoadPicture(App.Path & "\Flags\" & country(ctr) & ".gif")
                Flag = InputBox("What Country's Flag is This? Include " & "'_'" & " in Place of Spaces" & "Type " & " 'Quit'" & " to Quit")
                        'The Flag variable is the word that the user enters
                    If Flag = "Quit" Then   'gives the user the freedom to quit
                        GoTo endloop    'this stops the loop when the user decides to quit
                    ElseIf Flag = country(ctr) Then
                        MsgBox "You're Good!"
                            Score = Score + 1   'Begins to add up the number of correct answers
                    Else: MsgBox ("Incorrect.  The Country is " & country(ctr))
                    End If
Loop
endloop:
FinalScore = Score / ctr    'The final score from the game is calculated
MsgBox ("You're Score is " & FormatPercent(FinalScore, 0))  'The calculated Final Score is then formatted into a percentage

Close   'Close the array

End Sub
