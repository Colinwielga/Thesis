VERSION 5.00
Begin VB.Form frmSurvey2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Survey 2"
   ClientHeight    =   10290
   ClientLeft      =   7170
   ClientTop       =   3375
   ClientWidth     =   12480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   12480
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3240
      Picture         =   "Form_Survey2.frx":0000
      ScaleHeight     =   3615
      ScaleWidth      =   5295
      TabIndex        =   11
      Top             =   120
      Width           =   5295
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
      TabIndex        =   10
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdFrmSurvey3 
      BackColor       =   &H000080FF&
      Caption         =   "On to round 3!"
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
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picResults8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   8
      Top             =   8880
      Width           =   4335
   End
   Begin VB.PictureBox picResults7 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   7
      Top             =   7560
      Width           =   4335
   End
   Begin VB.PictureBox picResults6 
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton cmdHoles 
      BackColor       =   &H8000000D&
      Caption         =   "Name something you wear even if it has a hole in it"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   26.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   12
      Top             =   6120
      Width           =   1455
   End
End
Attribute VB_Name = "frmSurvey2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Family Feud
'frmSurvey2
'Colin Hall and Andre Blaine
'March 21
'The objective of this form is to receive an answer from the user and display the answer in the correct box
'This button hides the current form and shows the next form
Private Sub cmdFrmSurvey3_Click()
    frmSurvey2.Hide
    FrmSurvey3.Show
    MsgBox "Now the points are tripled!", , "Jackpot!" 'Informs the reader the points are tripled
End Sub

Private Sub cmdHoles_Click()
'This button receives an answer from the user and displays if it is correct or not
'The data used is read from a file
'The data consists of pairs of an answer and its value
'The data is read into two parallel arrays and the array containing the answers
'is searched to find an answer that matches the answer given by the user
'When the anwser is found the search stops and displays the answer and corresponding value
'into the appropriate output box

'Declare the Variables
Dim Holes(1 To 10) As String, Value(1 To 10) As Integer, CTR As Integer, Answer As String, X As Integer
Dim Found As Boolean, Strikes As Integer, Total As Integer, Remaining As Integer
picResultsSum.Print Sum

'Open the data file
Open App.Path & "\holes.txt" For Input As #1

Do While Not EOF(1)     'This loop reads data from a file into two arrays
    CTR = CTR + 1       'Increment the counter
    Input #1, Holes(CTR), Value(CTR)      'Get the next answer and value from the user
Loop

Do While Strikes < 3 And Total < 6      'Repeats the search until either all 6 answers are found, or the user has guessed three wrong answers
Answer = InputBox("Enter your answer in all lower case letters please", "Answer!")  'Get an answer from the user to use in the search
Found = False
    Do While ((Not Found) And (X < CTR))    'Searches the array until the answer is found or til the end of list
        X = X + 1
        If Answer = Holes(X) Then     'Compare every value on the list with the answer given by the user
            Found = True
                Select Case X       'Prints the answer in the correct box
                    Case Is = 1
                        picResults1.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults1.Print Holes(X), Value(X)
                    Case Is = 2
                        picResults2.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults2.Print Holes(X), Value(X)
                    Case Is = 3
                        picResults3.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults3.Print Holes(X), Value(X)
                    Case Is = 4
                        picResults4.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults4.Print Holes(X), Value(X)
                    Case Is = 5
                        picResults5.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults5.Print Holes(X), Value(X)
                    Case Is = 6
                        picResults6.Picture = LoadPicture(App.Path & "\white.jpg")
                        picResults6.Print Holes(X), Value(X)
                End Select
            Total = Total + 1       'They got an answer right, increase the total by one
            Sum = Value(X) * 2 + Sum    'Double the value and add it to the sum because points are doubled
            picResultsSum.Cls       'Clear the previous value of the sum
            picResultsSum.Print Sum     'Print the new value of the sum
        End If
    Loop

'This increments the strikes by one if the answer from the user is not in the list
If (Not Found) Then
    Strikes = Strikes + 1
    Remaining = 3 - Strikes
    MsgBox "Sorry, but that is not one of the answers! You only have " & Remaining & " remaining.", , "Sorry"
    
End If
X = 0       'Resets the value of X so the search will start from the first answer in the list
Loop

'Shows when you have three strikes
If Strikes = 3 Then
    MsgBox "You got three strikes :(", , "Failure"
    MsgBox "Let's see what you missed, then it's onto the next round!", , "Hooray"
End If

'Shows when you got all the answers
If Total = 6 Then
    MsgBox "Good Work! You got all the answers right! On to the next round!", , "Great Success!"
End If

'This shows all the anwsers, both missed and not
picResults1.Cls
picResults1.Picture = LoadPicture(App.Path & "\white.jpg")
picResults1.Print Holes(1), Value(1)
picResults2.Cls
picResults2.Picture = LoadPicture(App.Path & "\white.jpg")
picResults2.Print Holes(2), Value(2)
picResults3.Cls
picResults3.Picture = LoadPicture(App.Path & "\white.jpg")
picResults3.Print Holes(3), Value(3)
picResults4.Cls
picResults4.Picture = LoadPicture(App.Path & "\white.jpg")
picResults4.Print Holes(4), Value(4)
picResults5.Cls
picResults5.Picture = LoadPicture(App.Path & "\white.jpg")
picResults5.Print Holes(5), Value(5)
picResults6.Cls
picResults6.Picture = LoadPicture(App.Path & "\white.jpg")
picResults6.Print Holes(6), Value(6)
Close #1
cmdFrmSurvey3.Visible = True    'Displays the button to go onto the next round
cmdHoles.Enabled = False      'Disables the survey button
End Sub

Private Sub Form_Load()
picResults1.Picture = LoadPicture(App.Path & "\1.jpg")
picResults2.Picture = LoadPicture(App.Path & "\2.jpg")
picResults3.Picture = LoadPicture(App.Path & "\3.jpg")
picResults4.Picture = LoadPicture(App.Path & "\4.jpg")
picResults5.Picture = LoadPicture(App.Path & "\5.jpg")
picResults6.Picture = LoadPicture(App.Path & "\6.jpg")
picResults7.Picture = LoadPicture(App.Path & "\blank.jpg")
picResults8.Picture = LoadPicture(App.Path & "\blank.jpg")

End Sub
