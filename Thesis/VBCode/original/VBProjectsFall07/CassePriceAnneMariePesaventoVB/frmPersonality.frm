VERSION 5.00
Begin VB.Form frmPersonality 
   BackColor       =   &H003D30AD&
   Caption         =   "Personality Quiz"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14130
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H003D30AD&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBookStore 
      Caption         =   "Help Belle Buy A Book"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7440
      TabIndex        =   5
      Top             =   8040
      Width           =   2895
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "See Results History"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      TabIndex        =   4
      Top             =   8040
      Width           =   3135
   End
   Begin VB.CommandButton cmdTakeQuiz 
      Caption         =   "Take Quiz"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0084C11E&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10920
      MaskColor       =   &H0084C11E&
      TabIndex        =   1
      Top             =   8040
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   2880
      Picture         =   "frmPersonality.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   1920
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackColor       =   &H003D30AD&
      Caption         =   "Which Beauty and the Beast          Character are you?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   1815
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frmPersonality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBookStore_Click()
'allow user to move from main form to BookStore form
frmPersonality.Hide
frmBookStore.Show

End Sub

Private Sub cmdHistory_Click()
'allow user to move from the main form to Results History form
frmPersonality.Hide
frmHistory.Show

End Sub

Private Sub cmdQuit_Click()
'allow user to exit the program
End
End Sub

Private Sub cmdTakeQuiz_Click()
'moves user from main form to quiz question 1
frmPersonality.Hide
frmQuestion1.Show
'User enters name and it is displayed with a welcome message in a message box
Player = InputBox("Please enter your first name")
MsgBox "Welcome " & Player & " let's figure out which character you're most like!"
End Sub

