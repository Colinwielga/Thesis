VERSION 5.00
Begin VB.Form frmTrivia 
   Caption         =   "Trivia"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   Picture         =   "frmTrivia.frx":0000
   ScaleHeight     =   9030
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGostore 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go to Our Store"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go Back to Homepage"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   3735
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   1200
      ScaleHeight     =   3435
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   5400
      Width           =   6255
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0000FFFF&
      Caption         =   "Start Trivia"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   4095
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGostore_Click()
frmProducts.Show
frmTrivia.Hide
End Sub

Private Sub cmdQuit_Click()
frmHome.Show
frmTrivia.Hide
End Sub

Private Sub cmdStart_Click()
Dim qteam As String, ateam As String, qplayers As Integer, aplayers As Integer, qstrikes As Integer, astrikes As Integer, qbonds As String, abonds As String, qsox As String, asox As String, qleagues As String, aleagues As String, qchamps As Single, achamps As Single, qdiamond As Integer, adiamond As Integer
'Use Select Case to check if the input from the text box and the list are the same using
'conditions.
    
    ateam = "Twins"
    aplayers = 9
    astrikes = 3
    abonds = "yes"
    asox = "Red"
    aleagues = "American"
    achamps = "1991"
    asong = "Crackerjacks"
 
    NumMatches = 0
    
  'Ask questions using an input from an input box.
    qteam = InputBox("What is the nickname of Minnesota's MLB team?", "Question 1")
   
    'Check to see if the answers match the questions.
    If LCase(ateam) = LCase(qteam) Then
        NumMatches = NumMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qplayers = InputBox("How many players players play in the field at one time?", "Question 2")
    If aplayers = qplayers Then
        NumMatches = NumMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qstrikes = InputBox("How many strikes does a batter get?", "Question 3")
    If astrikes = qstrikes Then
        NumMatches = NumMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qbonds = InputBox("Does Barry Bonds use steriods?", "Question 4")
     If LCase(abonds) = LCase(qbonds) Then
        NumMatches = NumMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qsox = InputBox("What color Sox play in Boston ?", "Question 5")
    If LCase(asox) = LCase(qsox) Then
        NumMatches = NumMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qleagues = InputBox("There are two leagues; the National and what?", "Question 6")
    If LCase(aleagues) = LCase(qleagues) Then
        NumMatches = NumMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qchamps = InputBox("The Minnesota Twins were World Champions in 1987 and?", "Question 7")
    If LCase(achamps) = LCase(qchamps) Then
        NumMatches = NumMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    qsong = InputBox("Finish this song,Buy me some peanuts and ......", "Question 8")
    If LCase(asong) = LCase(qsong) Then
        NumMatches = NumMatches + 1
        MsgBox "Correct!", , "Answer!"
    End If
    
    picResults.Cls
    
    Select Case NumMatches
        Case 0
            picResults.Print "Are you from Mars!"
        Case 1
            picResults.Print "Baseball is our national pastime .  One right??? You know nothing!!"
        Case 2
            picResults.Print "You need to brush up on your baseball!"
        Case 3
            picResults.Print "Not very good, maybe you should try again!"
        Case 4
            picResults.Print "Could have done better!"
        Case 5
            picResults.Print "Half right, if you were in school you would get an F!!!"
        Case 6
            picResults.Print "Pretty good, someone watches baseball"
        Case 7
            picResults.Print "Wow, way to go!"
        Case 8
            picResults.Print "You are a basball GOD!!!"
        Case Else
            picResults.Print "You call yourself a baseball fan!"
    End Select
    'If more than 70% of their answers are correct, we will give them
    '10% off in our store with a message box.
    If NumMatches >= 7 Then
        MsgBox ("You have a 10 % discount waiting for you in our store!")
    End If
End Sub
