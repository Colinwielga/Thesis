VERSION 5.00
Begin VB.Form DateIntro 
   BackColor       =   &H008080FF&
   Caption         =   "Intro"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBudget 
      BackColor       =   &H008080FF&
      Height          =   375
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H008080FF&
      Height          =   375
      Left            =   6240
      ScaleHeight     =   315
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdCustomize 
      Caption         =   "Customize your experience"
      Height          =   615
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox picDate 
      Height          =   5535
      Left            =   480
      Picture         =   "StacyWurmIntro.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue on my Date"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label intro 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   $"StacyWurmIntro.frx":7B66
      Height          =   855
      Left            =   5640
      TabIndex        =   8
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Joke 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "(We all hope it goes well and that you don't end up like this guy!!)     <<-------------------"
      Height          =   855
      Left            =   6360
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label PDate 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Perfect Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6000
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label IntroParagraph 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   $"StacyWurmIntro.frx":7C01
      Height          =   855
      Left            =   5760
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "DateIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project Name: Date Chooser (Wurm, Stacy - VB Project)
' Form Name: DateIntro (StacyWurmIntro.frm)
' Author: Stacy Wurm
' Date Written: Sunday, March 7th, 2004
' Purpose of this Form: ' This project allows the user to choose their ideal date
                        ' They can go through the different objects
                        ' It also allows them to decide their own budget
                        ' It will tell how much is being spent and see if the user goes over budget
                        ' This form gets intitial information and budget

Private Sub cmdContinue_Click()
' moves the user on to the next form
DateIntro.Hide
GiftToGive.Show
End Sub

Private Sub cmdCustomize_Click()
' Get the name and budget from the user
UserName = InputBox("Please enter your name")
Budget = InputBox("How much would you like to spend at the most?")
picResults.Print Tab(10); UserName
picBudget.Print Tab(5); Budget
Joke.Visible = True
PDate.Visible = True
IntroParagraph.Visible = True
cmdContinue.Enabled = True
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
' Welcomes the user to the program
MsgBox "Welcome!!  Just wanted to welcome you to the fun and exciting program!!  Please click OK to continue!!", , "Welcome"
End Sub
