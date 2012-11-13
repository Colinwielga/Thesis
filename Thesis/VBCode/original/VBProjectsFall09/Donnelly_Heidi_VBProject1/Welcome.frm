VERSION 5.00
Begin VB.Form FrmWelcome 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetStarted 
      BackColor       =   &H0080C0FF&
      Caption         =   "Let's Get Started!"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "FrmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Bennie Health Project
    'FrmWelcome
    'Heidi Donnelly
    'Written: 10/5
    'The purpose of this form is to welcome the user and retrieve the user's name for display throughout the project.
    'The overall purpose of this project is to allow for college-aged women (Bennies) to look into their health and see where they stand.
    'There are three areas they can look into: Nutrition, Exercise, and Everything Else
    'Each area is broken up into categories or areas of analysis. They can look into how well they eat, how well they exercise, and ultimately how well they take care of themselves.
    'Throughout the whole project, they are able discover ways to improve their health if necessary.
    'It is a personal analysis tool as well as provider of health information.

Private Sub cmdGetStarted_Click()
'Retrieves and stores UserName in module/public
    UserName = InputBox("What's your name?", "Welcome!")
'hides Welcome page from user and shows main page to user
    FrmWelcome.Hide
    FrmMain.Show
'display greeting on Main Page
    MsgBox ("Hello, ") & UserName & ("! Let's find out how healthy you are! :)")
End Sub
