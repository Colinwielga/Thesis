VERSION 5.00
Begin VB.Form Main_Page 
   BackColor       =   &H0000FF00&
   Caption         =   "Home"
   ClientHeight    =   6465
   ClientLeft      =   6330
   ClientTop       =   3795
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8100
   Begin VB.CommandButton cmdSnacks 
      BackColor       =   &H0080FFFF&
      Caption         =   "Snacks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdAmIHealthy 
      BackColor       =   &H00FFFF00&
      Caption         =   "Am I Healthy?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdDinner 
      BackColor       =   &H00004080&
      Caption         =   "Dinner"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdLunch 
      BackColor       =   &H00008000&
      Caption         =   "Lunch"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdBreakfast 
      BackColor       =   &H000080FF&
      Caption         =   "Breakfast"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H0000FF00&
      Caption         =   $"Main_Page.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label lblSitationOfPicture 
      Caption         =   "http://www.mypyramid.gov/"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Please pick your meal:"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3315
      Left            =   240
      Picture         =   "Main_Page.frx":00F0
      Top             =   1320
      Width           =   4245
   End
End
Attribute VB_Name = "Main_Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main Page Code
'Calorie Counter
'By: Ryan Sutherland and Tara Johnson
'3-12-08
'This is a project that takes what you eat everyday for your three meals and calculates calories.
'It contains a module so that the variables that store the number of calories can be used on more
'than one form.  It allows the use on several forms, which we have.
Private Sub cmdAmIHealthy_Click()
Am_I_Healthy.Show      'this is used to move from one form to the other.
Main_Page.Hide
End Sub
'This moves the user from the main page to breakfast.
Private Sub cmdBreakfast_Click()
Main_Page.Hide
breakfast.Show
cmdAmIHealthy.Enabled = True        'allows the user to now acces the "Am I healthy" button.
End Sub
'This moves the user from the main page to dinner.
Private Sub cmdDinner_Click()
Main_Page.Hide
Dinner.Show
cmdAmIHealthy.Enabled = True        'allows the user to now acces the "Am I healthy" button.
End Sub
'This moves the user from the main page to lunch.
Private Sub cmdLunch_Click()
Main_Page.Hide
Lunch.Show
cmdAmIHealthy.Enabled = True        'allows the user to now acces the "Am I healthy" button.
End Sub
'Ends the project
Private Sub cmdQuit_Click()
End
End Sub
'This moves the user from the main page to snacks.
Private Sub cmdSnacks_Click()
Main_Page.Hide
Snacks.Show
cmdAmIHealthy.Enabled = True        'allows the user to now acces the "Am I healthy" button.
End Sub

