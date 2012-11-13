VERSION 5.00
Begin VB.Form frmOpening 
   BackColor       =   &H00FF8080&
   Caption         =   "Hello"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHomeGym 
      Height          =   3855
      Left            =   1800
      Picture         =   "frmOpening.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   6315
      TabIndex        =   3
      Top             =   1920
      Width           =   6375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFF00&
      Caption         =   "I Don't Want to Build Right Now"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5520
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   3495
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFF00&
      Caption         =   "Start Shopping!!!!"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Build Your Own Home Gym!"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1515
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Option Explicit makes the user declare all of their variables
'Project Name: Build Your Own Home Gym
'Form Name: frmOpening
'Author: Michelle Pickle
'Date Written: March 12th 2009
'The purpose of this form is the allow the user to begin the program and start building their very own home gym
'The overall purpose of this project is to allow the user to build their own home gym and recieve an estimated cost of builing it.
    'Additionally, the user is able to estimate the cost of belonging to a gym and/or purchase a membership to a gym.


Private Sub cmdQuit_Click()
'when the user clicks this button, the user will exit the program
    End
End Sub
Private Sub cmdStart_Click()
'declares the variable
Dim Name As String
'input box, which asks the individual his/her name and then stores it as "name"
    Name = InputBox("Welcome, Please Enter Your Name", "Welcome")
'a pop up message box welcomes the user by inserting his/her name after the "Welcome"
    MsgBox "Welcome " & Name, , "Welcome"
'after the message box, the opening screen disappears and the next screen is showen
    frmOpening.Hide
    frmHandHeld.Show
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading.
'this code discovered from Cassie Scherer and Jordan Schmaltz project of developing a vacation

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub

