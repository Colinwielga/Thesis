VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00FF0000&
   Caption         =   "Road To The Final Four"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdGetStarted 
      BackColor       =   &H000080FF&
      Caption         =   "CLICK TO GET STARTED!"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CREATED BY:  TJ ORTMANN AND RYAN KETTENACKER"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   3
      Top             =   7320
      Width           =   5415
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   720
      Top             =   0
      Width           =   6735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3030
      Left            =   2520
      Picture         =   "frmIntro.frx":0000
      Top             =   2040
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "WELCOME TO THE 2007 NCAA TOURNAMENT LETS GET STARTED AND FILL OUT YOUR BRACKET!"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Object of this program is to incorporate a scoring system directly in a
'March Madness Bracket for the Mens NCAA Basketball Tournament.

'with this form we wanted to give the user a quick intro page stating what our project is dealing with
'and to get the users name
Private Sub cmdGetStarted_Click()
    'Ask the user what there name is to use throughout the program
    User = InputBox("Please Enter Your Name", "Name")
    'will bring user to next form
    frmIntro.Hide
    frmInstructions.Show
End Sub

Private Sub cmdQuit_Click()
    'end program
    End
End Sub


