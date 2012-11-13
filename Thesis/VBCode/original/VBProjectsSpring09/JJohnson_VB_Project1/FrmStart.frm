VERSION 5.00
Begin VB.Form FrmStart 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   13740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "End Tour/Get to the airport"
      Height          =   1095
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton Cmdmatinee 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go to the Theatre"
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton CmdSoundSport 
      BackColor       =   &H000000FF&
      Caption         =   "Visit Yankee Stadium"
      Height          =   1095
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton CmdSights 
      BackColor       =   &H0000FF00&
      Caption         =   "Do some Sightseeing"
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label LblNYC1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "You have a layover in NYC and five hours before your plane takes off.  What do you want to do?"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   12015
   End
   Begin VB.Image Image1 
      Height          =   5925
      Left            =   600
      Picture         =   "FrmStart.frx":0000
      Top             =   960
      Width           =   12000
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Things to do in NYC
'Form Name: frmStart
'Author: Jake Johnson
'Date Written: 3/23/09
'Objective: This is the starting form for the project.
'The user gets to pick one of three different types of areas, sightseeing, sound/sports, and tastes/drink.
'Each of the areas will have different things to do with information available about their choices.

'goes to theatre form and options
Private Sub Cmdmatinee_Click()
FrmStart.Hide
frmTheatre.Show
End Sub

'Quits project
Private Sub CmdQuit_Click()
End
End Sub

'Goes to Sightseeing form and options
Private Sub CmdSights_Click()
FrmStart.Hide
FrmSightseeing.Show
End Sub

'goes to yankee form and options
Private Sub CmdSoundSport_Click()
FrmStart.Hide
FrmYankee.Show
End Sub
