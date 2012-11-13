VERSION 5.00
Begin VB.Form frmSurvey1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Survey 1"
   ClientHeight    =   10290
   ClientLeft      =   7170
   ClientTop       =   3210
   ClientWidth     =   12480
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   12480
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   3240
      Picture         =   "Form_Survey1.frx":0000
      ScaleHeight     =   3615
      ScaleWidth      =   5415
      TabIndex        =   12
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton cmdFrmSurvey2 
      BackColor       =   &H000080FF&
      Caption         =   "On to round 2!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picResultsSum 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      ScaleHeight     =   975
      ScaleWidth      =   1815
      TabIndex        =   9
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdPet 
      BackColor       =   &H8000000D&
      Caption         =   "Name a common household pet"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   26.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   2895
   End
   Begin VB.PictureBox picResults8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7440
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   7
      Top             =   8880
      Width           =   4335
   End
   Begin VB.PictureBox picResults7 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7440
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   6
      Top             =   7560
      Width           =   4335
   End
   Begin VB.PictureBox picResults6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7440
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   5
      Top             =   6240
      Width           =   4335
   End
   Begin VB.PictureBox picResults5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   4
      Top             =   4920
      Width           =   4335
   End
   Begin VB.PictureBox picResults4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   3
      Top             =   8880
      Width           =   4335
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   2
      Top             =   7560
      Width           =   4335
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   1
      Top             =   6240
      Width           =   4335
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H80000009&
      Caption         =   "Total points"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   6240
      Width           =   1455
   End
End
Attribute VB_Name = "frmSurvey1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Family Feud
'frmSurvey1
'Colin Hall and Andre Blaine
'March 21
'The objective of this form is to receive an answer from the user and display the answer in the correct box
Private Sub cmdFrmSurvey2_Click()
'This button hides the current form and shows the next form
    frmSurvey1.Hide
    frmSurvey2.Show
    MsgBox "Now the points are doubled!", , "Awesome!" 'Informs the reader the points are doubled
End Sub

Private Sub cmdPet_Click()
'This button receives an answer from the user and displays if it is correct or not
'The data used is read from a file
'The data consists of pairs of an answer and its value
'The data is read into two parallel arrays and the array containing the answers
'is searched to find an answer that matches the answer given by the user
'When the anwser is found the search stops and displays the answer and corresponding value
'into the appropriate output box

'Declare the variables
Dim Pet(1 To 10) As String, Value(1 To 10) As Integer, CTR As Integer, Answer As String, X As Integer
Dim Found As Boolean, Strikes As Integer, Total As Integer, Remaining As Integer


'Open the data file
Open App.Path & "\pets.txt" For Input As #1

Do While Not EOF(1)     'This loop reads data from a file into two arrays
    CTR = CTR + 1       'Increment the counter
    Input #1, Pet(CTR), Value(CTR)      'Get the next answer and value from the user
Loop

Do While Strikes < 3 And Total < 5      'Repeats the search until either all five answers are found, or the user has guessed three wrong answers
Answer = InputBox("Enter your answer in all lower case letters please", "Answer!")  'Get an answer from the user to use in the search
Found = False
    Do While ((Not Found) And (X < CTR))    'Searches the array until the answer is found or til the end of list
        X = X + 1
        If Answer = Pet(X) Then     'Compare every value on the list with the answer given by the user
            Found = True
                Select Case X       'Prints the answer in the correct box
                    Case Is = 1
                        picResults1.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults1.Print Pet(X), Value(X)
                    Case Is = 2
                        picResults2.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults2.Print Pet(X), Value(X)
                    Case Is = 3
                        picResults3.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults3.Print Pet(X), Value(X)
                    Case Is = 4
                        picResults4.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults4.Print Pet(X), Value(X)
                    Case Is = 5
                        picResults5.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults5.Print Pet(X), Value(X)
                End Select
            Total = Total + 1       'They got an answer right, increase the total by one
            Sum = Value(X) + Sum    'Add the value to the sum
            picResultsSum.Cls       'Clear the previous value of the sum
            picResultsSum.Print Sum     'Print the new value of the sum
        End If
    Loop

'This increments the strikes by one if the answer from the user is not in the list
If (Not Found) Then
    Strikes = Strikes + 1   'Increment strikes
    Remaining = 3 - Strikes
    MsgBox "Sorry, but that is not one of the answers! You have " & Remaining & " remaining.", , "Sorry" 'Tells the user how many strikes are left
    
End If
X = 0       'Resets the value of X so the search will start from the first answer in the list
Loop

'Shows when you have three strikes
If Strikes = 3 Then
    MsgBox "You got three strikes :(", , "Failure"
    MsgBox "Let's see what you missed, then it's onto the next round!", , "Hooray"
End If

'Shows when you got all the answers
If Total = 5 Then
    MsgBox "Good Work! You got all the answers right! On to the next round!", , "Great Success!"
End If

'This shows all the anwsers, both missed and not
picResults1.Cls
picResults1.Picture = LoadPicture(App.Path & "\white.jpg")
picResults1.Print Pet(1), Value(1)
picResults2.Cls
picResults2.Picture = LoadPicture(App.Path & "\white.jpg")
picResults2.Print Pet(2), Value(2)
picResults3.Cls
picResults3.Picture = LoadPicture(App.Path & "\white.jpg")
picResults3.Print Pet(3), Value(3)
picResults4.Cls
picResults4.Picture = LoadPicture(App.Path & "\white.jpg")
picResults4.Print Pet(4), Value(4)
picResults5.Cls
picResults5.Picture = LoadPicture(App.Path & "\white.jpg")
picResults5.Print Pet(5), Value(5)
Close #1
cmdFrmSurvey2.Visible = True    'Displays the button to go onto the next round
cmdPet.Enabled = False      'Disables the survey button
End Sub

Private Sub Form_Load()
picResults1.Picture = LoadPicture(App.Path & "\1.jpg")
picResults2.Picture = LoadPicture(App.Path & "\2.jpg")
picResults3.Picture = LoadPicture(App.Path & "\3.jpg")
picResults4.Picture = LoadPicture(App.Path & "\4.jpg")
picResults5.Picture = LoadPicture(App.Path & "\5.jpg")
picResults6.Picture = LoadPicture(App.Path & "\blank.jpg")
picResults7.Picture = LoadPicture(App.Path & "\blank.jpg")
picResults8.Picture = LoadPicture(App.Path & "\blank.jpg")

End Sub
