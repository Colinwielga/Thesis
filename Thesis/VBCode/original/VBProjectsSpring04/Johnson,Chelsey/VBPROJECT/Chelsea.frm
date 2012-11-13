VERSION 5.00
Begin VB.Form Chelsea 
   BackColor       =   &H000080FF&
   Caption         =   "Chelsea"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
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
      Height          =   1095
      Left            =   16200
      TabIndex        =   12
      Top             =   10320
      Width           =   1815
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14160
      TabIndex        =   11
      Top             =   10320
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   10
      Text            =   "Did we choose the same 1st choice?"
      Top             =   8160
      Width           =   4215
   End
   Begin VB.PictureBox picResults 
      Height          =   1935
      Left            =   7800
      ScaleHeight     =   1875
      ScaleWidth      =   4395
      TabIndex        =   9
      Top             =   7200
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12480
      TabIndex        =   8
      Text            =   "Battersea Bridge"
      Top             =   4800
      Width           =   1935
   End
   Begin VB.OptionButton optbattersea 
      Caption         =   "Option1"
      Height          =   255
      Left            =   12120
      TabIndex        =   7
      Top             =   4800
      Width           =   255
   End
   Begin VB.PictureBox picBattersea 
      Height          =   3015
      Left            =   10800
      Picture         =   "Chelsea.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   6075
      TabIndex        =   6
      Top             =   1680
      Width           =   6135
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4440
      TabIndex        =   5
      Text            =   "Albert Bridge"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.OptionButton optalbertbridge 
      Caption         =   "Option1"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   4440
      Width           =   255
   End
   Begin VB.PictureBox picalbertbridge 
      Height          =   2295
      Left            =   3960
      Picture         =   "Chelsea.frx":63CF
      ScaleHeight     =   2235
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Text            =   "Then vote on which site is your favorite."
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6360
      TabIndex        =   1
      Text            =   "Click on each picture to learn more about that famous site."
      Top             =   600
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Text            =   "Each picture represents a famous site within Chelsea district."
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   11880
      Width           =   1935
   End
End
Attribute VB_Name = "Chelsea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: Chelsea (Chelsea.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: This form is to let the user become familiar with Albert Bridge and Battersea Bridge.
                'The are then able to vote on which one they choose to be their favorite, and they
                'can then see if it is the same as my first choice
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdgoback_Click()
'Brings the user back to the first page, with the map of London, this is so they can choose a new district
Chelsea.Hide
MapLondon.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub optalbertbridge_Click()
'If they choose this option button they are informed that they have choosen the same site as I had
picResults.Cls
picResults.Print "Congradulations!"
picResults.Print "Albert Bridge is also my favorite site in Chelsea."
End Sub


Private Sub optbattersea_Click()
'If they choose this option button they are informed that they had not choosen the same site as I had
picResults.Cls
picResults.Print "Sorry, we do not have the same 1st choice."
picResults.Print "Battersea is my 2nd choice."
End Sub

Private Sub picalbertbridge_Click()
'By clicking on the picture the user is informed of the history of the Albert Bridge
MsgBox "A 1864 Act of Parliament authorised the construction of a bridge but there were long delays before it was opened to traffic in 1873.", , "Albert Bridge"
MsgBox "The bridge was designed by Roland Mason Ordish, Albert Bridge was originally a cantilever bridge, with each half of the bridge being supported by bars radiating out from the top of its supporting towers.", , "Albert Bridge"
MsgBox "The 710 ft long bridge was made up of two side spans of 155 ft and a centre of 400 ft.  The roadway was 41 ft in width.", , "Albert Bridge"
End Sub

Private Sub picBattersea_Click()
'By clicking on the picture the user is informed of the history of the Battersea Bridge
MsgBox "The original Battersea Bridge, when built in 1772, was made of wood to the design of Henry Holland, and was the only bridge between Westminster and Putney.", , "Battersea Bridge"
MsgBox "It had the effect of transforming the village of Chelsea into a town. But it was very dangerous and my boats collided with it.  It was eventually closed.", , "Battersea Bridge"
MsgBox "In 1890 the present five arched, cast iron bridge, designed by Bazalgette was opened.", , "Battersea Bridge"
End Sub
