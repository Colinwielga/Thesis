VERSION 5.00
Begin VB.Form frmhome 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdplayers 
      Caption         =   "Is he on the Vikings?"
      Height          =   975
      Left            =   5040
      TabIndex        =   4
      Top             =   5400
      Width           =   2655
   End
   Begin VB.CommandButton cmdroster 
      Caption         =   "View the Vikings' Schedule"
      Height          =   975
      Left            =   5040
      TabIndex        =   3
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmdtrivia 
      Caption         =   "Take The Trivia Challenge"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmdorder 
      Caption         =   "Order Vikings Gear"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label lblhome 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Minnesota Vikings Fan Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   -120
      Picture         =   "VBproject.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7980
   End
End
Attribute VB_Name = "frmhome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Title: Minnesota Vikings Fan Page

'Form Name: Home

'Written by Jim Zrust

'Date: November 1, 2006

'Form Objective:'the front page was only used to take the user to different forms where the bulk of the program lays

'Objective of program:  When I started writing this I wanted to create a program having
'something to do with the Minnesota Vikings that would be interesting to look through.
'I wanted to give the user various options that are presented on the front page that they
'can choose to complete.  I ended up deciding to have four different options which were:
'search the roster, view the schedule in different ways, shop at the Vikings "Team Store",
'and give the user an opportunity to test their Vikings knowledge through a trivia game.

Private Sub cmdorder_Click() 'takes the user to new form
frmhome.Hide
frmorder.Show
End Sub

Private Sub cmdplayers_Click()
frmhome.Hide
frmplayers.Show
End Sub

Private Sub cmdroster_Click()
frmhome.Hide
frmschedule.Show
End Sub

Private Sub cmdtrivia_Click()
frmhome.Hide
frmtrivia.Show
End Sub




