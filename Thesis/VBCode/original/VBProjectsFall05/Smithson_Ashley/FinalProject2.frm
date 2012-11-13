VERSION 5.00
Begin VB.Form FinalProject2 
   BackColor       =   &H00404000&
   Caption         =   "FinalProject2"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdinfo 
      Caption         =   "Useful Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   8
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton cmdplaces 
      BackColor       =   &H00E0E0E0&
      Caption         =   "The Places We Went"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton cmdactivites 
      BackColor       =   &H00E0E0E0&
      Caption         =   "What We Did"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdpeople 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Friends"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton cmdpo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Where We Lived"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   4095
      Left            =   840
      Picture         =   "FinalProject2.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label Ashley 
      BackColor       =   &H00404000&
      Caption         =   "Ashley K. Smithson"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Australia 
      BackColor       =   &H00404000&
      Caption         =   "Australia"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FinalProject2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Australia
'Formname: FinalProject2
'Author: Ashley Smithson
'Date: October 31, 2005
'Purpose of Project: To show others all you can do with a wonderful experience abroad
'Purpose of Form: starting block; directs you to other information
Option Explicit
Private Sub cmdactivites_Click()
FinalProject2.Hide 'brings you to the activities page
Activites.Show
End Sub

Private Sub cmdinfo_Click()
FinalProject2.Hide 'brings you to the usefull information page
UsefulInformation.Show
End Sub

Private Sub cmdpeople_Click()
FinalProject2.Hide 'brings you to the friends page
People.Show
End Sub

Private Sub cmdplaces_Click()
FinalProject2.Hide 'brings you to the places page
Places.Show
End Sub

Private Sub cmdpo_Click()
FinalProject2.Hide 'brings you to the PennOrient page
PennOrient.Show
End Sub


Private Sub cmdquit_Click()
End 'quits the program
End
End Sub
