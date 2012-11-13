VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H000000FF&
   Caption         =   "Welcome to SportsCalculator! Helping people choose sports for 2 weeks running!"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   10200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrOne 
      Interval        =   5000
      Left            =   4440
      Top             =   4800
   End
   Begin VB.Label lblNeed 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Need help choosing a past time?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7920
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lblClue 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Lazy... or Clueless?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label lblSport 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "I know I'm fast... and thats about it!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO THE SPORTS CALCULATOR!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   4665
      Left            =   2280
      Picture         =   "Form1.frx":0000
      Top             =   2040
      Width           =   6000
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrOne_Timer() 'delays the opening of the program so the user can read the data and look at he picture
    frmWelcome.Hide
    frmFirst.Show
    tmrOne = True
    tmrOne = False 'prevents the reoccurence of the timer
End Sub
