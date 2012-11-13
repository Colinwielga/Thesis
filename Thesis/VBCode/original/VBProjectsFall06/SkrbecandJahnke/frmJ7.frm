VERSION 5.00
Begin VB.Form frmCeleb400 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Celeb 400"
   ClientHeight    =   7260
   ClientLeft      =   7605
   ClientTop       =   5655
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   9495
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Celebrity Category"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1695
   End
   Begin VB.PictureBox picOutput2 
      Height          =   2055
      Left            =   5640
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox txtFriends 
      Height          =   1335
      Left            =   5640
      TabIndex        =   2
      Top             =   2520
      Width           =   3135
   End
   Begin VB.PictureBox picOutput 
      Height          =   3735
      Left            =   1320
      ScaleHeight     =   3675
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CommandButton cmdActors 
      BackColor       =   &H00FF8080&
      Caption         =   "Push Button to See List of Actors"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblC400 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmJ7.frx":0000
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmCeleb400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pos As Integer, Friends(1 To 6) As String           'Sets the dimension as string since we
'Jeopardy.(Jeopardy.vbp)                                'are using text in out text file (instead
'Form name: Celeb400; Form caption: Jeopardy            'of just numbers)
'Author: Skrbec and Jahnke
'Date written: October 29, 2006
'Form Objective: This is the Celeb400 form. It is used when the user clicks on the 400 point
'                question. It will open a text file and then display what is in it. The user
'                will then read what is in the picturebox and answer the question.


Private Sub cmdActors_Click()                       'Has user click in order to enter answer
pos = 0                                             'Sets counter to 0
Open App.Path & "\celebrity.txt" For Input As #1    'Opens the text file, labels as #1
    Do Until EOF(1)                                 'Reads the file until the end
        pos = pos + 1                               'Creates a counter
        Input #1, Friends(pos)                      'Stores the name at that point in the file
        picOutput.Print Friends(pos)                'Prints the name of what it just read
    Loop                                            'Loops the "Do Until" so it will print all of the names
Close #1                                            'Closes the text file
End Sub

Private Sub cmdReturn_Click()
frmCeleb400.Hide
frmCelebrities.Show
End Sub

Private Sub txtFriends_DblClick()                   'Has user double-click to submit answer
    If txtFriends.Text = "friends" Then
        picOutput2.Print "You are Correct!  Great Job!"
        picOutput2.Print "Please click the Return To"
        picOutput2.Print "Celebrities button!"
        Sum = Sum + 400
    Else
        picOutput2.Print "Wrong Answer. Please click the"
        picOutput2.Print "Return To Celebrities button!"
        Sum = Sum - 400
    End If
End Sub
