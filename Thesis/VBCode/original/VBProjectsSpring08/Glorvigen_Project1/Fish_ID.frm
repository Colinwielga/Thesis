VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0080C0FF&
   Caption         =   "Fish ID Game"
   ClientHeight    =   4545
   ClientLeft      =   4245
   ClientTop       =   3630
   ClientWidth     =   7320
   LinkTopic       =   "Form3"
   ScaleHeight     =   4545
   ScaleWidth      =   7320
   Visible         =   0   'False
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdresults 
      BackColor       =   &H00FF8080&
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to start the game!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4200
      Picture         =   "Fish_ID.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H008080FF&
      Caption         =   "Leave Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0080FF80&
      Caption         =   "Back To Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Name These Fish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Fisher
'Fish ID Game
'Eric Glorvigen
'Date= March 5
'this is the opening page of the fish id game which naviagtes from page to
'page and asks the user questions, it finally configures the results
'and tells the user how he or she did and what questions they got wrong


Private Sub cmdexit_Click()
    'exit back to main page
        form1.Show
        Form3.Hide
End Sub

Private Sub cmdleave_Click()
    'leave program
        End
End Sub

Private Sub cmdresults_Click()
 'tally the questions that were correct
 'this uses all global variables and a if then else statement
 'and a select case statement

    If questionctr <> 5 Then
    
        picoutput.Print "The Questions You missed were:"
        picoutput.Print "************************************"
        
        If wrongone = True Then
                picoutput.Print "#1"
        End If

        If wrongtwo = True Then
            picoutput.Print "#2"
        End If
    
        If wrongthree = True Then
            picoutput.Print "#3"
        End If
            
        If wrongfour = True Then
            picoutput.Print "#4"
        End If
         
        If wrongfive = True Then
            picoutput.Print "#5"
        End If
    End If
    
    picoutput.Print "*******************************"
  
      Select Case questionctr
        Case Is = 5
            picoutput.Print "YOU ARE A MASTER!"
        Case Is = 4
            picoutput.Print "Almost perfect!"
        Case Is = 3
            picoutput.Print "A little over half... Keep studying"
        Case Is = 2
            picoutput.Print "Sorry only two right, you must be the perch person!"
        Case Is = 1
            picoutput.Print "You need to fish more!"

    End Select
  
End Sub

Private Sub cmdstart_Click()
    'start quiz by moving to question 1
        Form4.Show
        Form3.Hide
End Sub


