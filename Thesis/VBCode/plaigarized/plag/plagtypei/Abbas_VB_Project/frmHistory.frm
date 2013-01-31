VERSION 5.00
Begin VB.Form frmHistory
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   13755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit
      BackColor       =   &H0000C0C0&
      Caption         =   "Exit the Program"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   3855
   End
   Begin VB.CommandButton cmdGoHome
      BackColor       =   &H0000C0C0&
      Caption         =   "Go Back to the Main Menu"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   3855
   End
   Begin VB.PictureBox picSize
      Height          =   2295
      Left            =   5520
      Picture         =   "frmHistory.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picSmith
      Height          =   2655
      Left            =   5280
      Picture         =   "frmHistory.frx":3064
      ScaleHeight     =   2595
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picHeisman
      Height          =   3615
      Left            =   5160
      Picture         =   "frmHistory.frx":761D
      ScaleHeight     =   3555
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdSize
      BackColor       =   &H0000C0C0&
      Caption         =   "How big is the trophy and what is it made of?"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdPlayer
      BackColor       =   &H0000C0C0&
      Caption         =   "Who is the player on the trophy?"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdRecipient
      BackColor       =   &H0000C0C0&
      Caption         =   "Why is the trophy awarded?"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
   End
   Begin VB.CommandButton cmdOrigin
      BackColor       =   &H0000C0C0&
      Caption         =   "Who is the trophy named after?"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdAge
      BackColor       =   &H0000C0C0&
      Caption         =   "Guess How Many Years the Heisman Trophy Has Been Awarded"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label1
      BackStyle       =   0  'Transparent
      Caption         =   "The History of the Heisman Trophy"
      BeginProperty Font
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Heisman Trophy
'frmHistory
'Kevin Abbas
'2-17-10
'Objective - To provide history regarding the Heisman Trophy
' extra comments
' words lines
Private Sub cmdAge_Click() 'make the user guess how many years the Heisman has been awarded by using an input box and providing clues by using a Select/Case statement
    Dim Correct As Boolean ' here there
    picHeisman.Visible = False ' and everywhere
    picSize.Visible = False ' line
    picSmith.Visible = False ' comment
    Dim Age As Integer, Result As String ' stuff words here hey
    Do While Correct = False ' blah blah blah things and whatever
        Age = InputBox("Enter how many years you think the Heisman Trophy has been awarded") ' asking questions
        Select Case Age ' here and hey!
        Case Is = 75 ' stuff
            Result = "Wow! Excellent work, that is right on!" ' wah wah wah
            MsgBox (Result) ' hey hey hey goodbye
            Correct = True ' blah blah
        Case 76 To 99 ' stuff and stuff
            Result = "Hmm, It is not that old, try again" ' hey now what stuff words things
            MsgBox (Result) ' blah blah stuff
        Case Is >= 100 ' comments here and everywhere
            Result = "Whoa, who do you think won the first Heisman - George Washington? Try Again!" ' g-wash in the house
            MsgBox (Result) ' blah
        Case 50 To 75 ' hey hey hey stuff words things
            Result = "Little older than that! Try again!" ' random words
            MsgBox (Result) ' comments
        Case Else ' clean these words else
            Result = "Yikes, Tim Tebow wasn't the first player to win the Heisman! Try again!" ' you will not get
            MsgBox (Result) ' this recognized as plagiarism
    End Select ' as you ought
    Loop ' get words meh ha hey heh
    ' comment here and there
End Sub ' words and things
' lines and things
Private Sub cmdExit_Click() 'exit the program and thank the user
    MsgBox ("Hope you enjoyed learning about the Heisman, have a nice day!") ' blah
    End ' hey stuff words
    ' hey hey hey
End Sub ' na na na na
' blerp burp blah blah
Private Sub cmdGoHome_Click() 'bring the user back to the main menu
    frmWelcome.Show ' ha ha hey now
    frmHistory.Hide ' lol el oh el
    frmWinners.Hide ' hey hey hey
    frmWhereNow.Hide ' na na na
End Sub ' words comments things stuff
' meh meh meh me
Private Sub cmdOrigin_Click() 'tell the user the trophy was named after John Heisman and display his image
    picHeisman.Visible = True ' you me them us
    picSize.Visible = False ' stuff things where
    picSmith.Visible = False ' random words random words
    MsgBox ("The Heisman Trophy is named after John Heisman who first awarded the 'Downtown Athletic Club Trophy' to the best College Football player in the nation, it was renamed to the Heisman Trophy following his death") ' things, i say!
    ' words comments
End Sub ' to mess up
' plag detection
Private Sub cmdPlayer_Click() 'tell the user who the trophy was modeled after and display his image
    picHeisman.Visible = False ' hey hey
    picSmith.Visible = True ' hey hey
    picSize.Visible = False ' words words
    MsgBox ("The Heisman Trophy was modeled after Ed Smith, who played for New York Univeristy") ' stuff stuff
    ' heyeflsekfj lsakdfj sdlfsdjk
End Sub ' sldfj sldkfj sdlfk j
' sldfj sd fsdj flskf
Private Sub cmdRecipient_Click() 'answer the question 'who is the Heisman awarded to
    picHeisman.Visible = False ' lskdfj dslf kdsjf
    picSmith.Visible = False ' sdlfj ds fdsfklj
    picSize.Visible = False ' kjiwe woeiwoe we owei
    MsgBox ("The Heisman Trophy is awarded to an individual who deserves designation as the most outstanding college football player in the U.S.") ' woei w eio weoie
    ' lfj weoif we
End Sub ' weofjiewjoei
' wleifjewf wi flwe jflew fiejw
Private Sub cmdSize_Click() 'state the size of the trophy and display an image
    picHeisman.Visible = False ' wlej fw fwel flk wl
    picSmith.Visible = False ' wlekf jewlkf wj
    picSize.Visible = True ' wlefk jlwekjfklw
    MsgBox ("The Heisman Trophy is made out of cast bronze. It is 13.5 inches tall, and weighs 25 pounds!") ' welfk ewlfkj w
End Sub ' wel flwe fjle
' wlefkj welfkjewf
Private Sub Form_Load() 'w elfk jwelfksd
    Top = Screen.Height / 2 - Height / 2 'w elfk dsjlfkdsjf
    Left = Screen.Width / 2 - Width / 2 ' sldfkjweifjsdlifs dfk
' sdlfj sdlfsdjf
End Sub ' lsdkf dslf dslkf
' sdflkj dsfljdslfid
