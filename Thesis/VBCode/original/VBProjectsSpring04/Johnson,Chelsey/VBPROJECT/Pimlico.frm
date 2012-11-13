VERSION 5.00
Begin VB.Form Pimlico 
   BackColor       =   &H0000FF00&
   Caption         =   "Pimlico"
   ClientHeight    =   12585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   ScaleHeight     =   12585
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      TabIndex        =   10
      Top             =   10800
      Width           =   1455
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Click Here to Return to the Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10920
      TabIndex        =   9
      Top             =   9600
      Width           =   2775
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Click Here learn new information on the site of your choice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   1560
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   10395
      TabIndex        =   7
      Top             =   7200
      Width           =   10455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   6
      Text            =   "2. Tate Gallery"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox picTate 
      Height          =   4215
      Left            =   5280
      Picture         =   "Pimlico.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   4395
      TabIndex        =   5
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox txtsite 
      Height          =   495
      Left            =   9960
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "1. Vauxhall Bridge"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.PictureBox picvauxhall 
      Height          =   1575
      Left            =   480
      Picture         =   "Pimlico.frx":7844
      ScaleHeight     =   1515
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Text            =   "Type in the number next to the picture to learn more about each site."
      Top             =   840
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Text            =   "These pictures are of famous sites found within the district of Pimlico."
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   11760
      Width           =   2535
   End
End
Attribute VB_Name = "Pimlico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: Pimlico (Pimlico.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of form: The purpose of this form is to let the user learn the history of Vauxhall Bridge and The Tate Gallery,
                    'by choosing the one they want to learn about first by entering a number.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdcompute_Click()
Dim Site As Integer
Site = txtsite.Text 'Getting Variable from the user
If Site = "1" Then
    picResults.Cls
    'Printing the results if the user had chosen number 1
    picResults.Print "Vauxhall Bridge is an iron bridge, of nine equal arches, over the Thames at Vauxhall, communicating with Millbank on the left bank of the river."
    picResults.Print "It was built from the designs of James Walker; commenced May 9th, 1811, and opened June 4th, 1816."
End If
If Site = "2" Then
    picResults.Cls
    'Printing the results if the user had chosen number 2
    picResults.Print "The Tate Gallery houses the national collection of modern Art."
End If
End Sub


Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Returns the user back to the Map of London Page, so they can choose a new district to look at.
Pimlico.Hide
MapLondon.Show

End Sub

