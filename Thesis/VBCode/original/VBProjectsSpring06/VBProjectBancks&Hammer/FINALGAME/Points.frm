VERSION 5.00
Begin VB.Form frmpoints 
   BackColor       =   &H0000FFFF&
   Caption         =   "Points"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScore 
      Caption         =   "Show Me My Final Score!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      Picture         =   "Points.frx":0000
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   2
      Top             =   6240
      Width           =   1455
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4920
      ScaleHeight     =   2595
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Label lblnames 
      BackStyle       =   0  'Transparent
      Caption         =   "by Lisa Hammer and Kate Bancks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   7440
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   5715
      Left            =   480
      Picture         =   "Points.frx":07CD
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   3960
   End
   Begin VB.Label lblpoints 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks For Playing CIRCUS FUN!! "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   7575
   End
End
Attribute VB_Name = "frmPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()                         'the purpose of frmpoints is to show the user their final score and to give relative feedback of their performance.
    End                                             'this button allows the user to quit the program
End Sub

Private Sub cmdScore_Click()                        'this button shows the final score and gives final feedback.
    picresults.Print N; " Your Final Score is:", C
        Select Case C
            Case 0 To 9
                picresults.Print "Great Start, Try Again!"
            Case 10 To 19
                picresults.Print "Excellent!"
            Case 20 - 30
                picresults.Print "UNBELIEVABLE JOB!"
        End Select
    
End Sub
