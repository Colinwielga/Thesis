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


Private Sub cmdAge_Click() 'make the user guess how many years the Heisman has been awarded by using an input box and providing clues by using a Select/Case statement
    Dim Correct As Boolean
    picHeisman.Visible = False
    picSize.Visible = False
    picSmith.Visible = False
    Dim Age As Integer, Result As String
    Do While Correct = False
        Age = InputBox("Enter how many years you think the Heisman Trophy has been awarded")
        Select Case Age
        Case Is = 75
            Result = "Wow! Excellent work, that is right on!"
            MsgBox (Result)
            Correct = True
        Case 76 To 99
            Result = "Hmm, It is not that old, try again"
            MsgBox (Result)
        Case Is >= 100
            Result = "Whoa, who do you think won the first Heisman - George Washington? Try Again!"
            MsgBox (Result)
        Case 50 To 75
            Result = "Little older than that! Try again!"
            MsgBox (Result)
        Case Else
            Result = "Yikes, Tim Tebow wasn't the first player to win the Heisman! Try again!"
            MsgBox (Result)
    End Select
    Loop
    
End Sub

Private Sub cmdExit_Click() 'exit the program and thank the user
    MsgBox ("Hope you enjoyed learning about the Heisman, have a nice day!")
    End
    
End Sub

Private Sub cmdGoHome_Click() 'bring the user back to the main menu
    frmWelcome.Show
    frmHistory.Hide
    frmWinners.Hide
    frmWhereNow.Hide
End Sub

Private Sub cmdOrigin_Click() 'tell the user the trophy was named after John Heisman and display his image
    picHeisman.Visible = True
    picSize.Visible = False
    picSmith.Visible = False
    MsgBox ("The Heisman Trophy is named after John Heisman who first awarded the 'Downtown Athletic Club Trophy' to the best College Football player in the nation, it was renamed to the Heisman Trophy following his death")
    
End Sub

Private Sub cmdPlayer_Click() 'tell the user who the trophy was modeled after and display his image
    picHeisman.Visible = False
    picSmith.Visible = True
    picSize.Visible = False
    MsgBox ("The Heisman Trophy was modeled after Ed Smith, who played for New York Univeristy")
    
End Sub

Private Sub cmdRecipient_Click() 'answer the question 'who is the Heisman awarded to
    picHeisman.Visible = False
    picSmith.Visible = False
    picSize.Visible = False
    MsgBox ("The Heisman Trophy is awarded to an individual who deserves designation as the most outstanding college football player in the U.S.")
    
End Sub

Private Sub cmdSize_Click() 'state the size of the trophy and display an image
    picHeisman.Visible = False
    picSmith.Visible = False
    picSize.Visible = True
    MsgBox ("The Heisman Trophy is made out of cast bronze. It is 13.5 inches tall, and weighs 25 pounds!")
End Sub

Private Sub Form_Load()
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
