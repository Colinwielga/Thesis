VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1(Opening Form)"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGlossary 
      Caption         =   "Glossary of Graffiti Terms"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   4
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   6240
      Width           =   2655
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00000000&
      Caption         =   "Begin Lesson"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmdNames 
      BackColor       =   &H00000000&
      Caption         =   "View List of My Favorite Graffiti Artists"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "~frank verdeja"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Intro to Graffiti"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Intro to Graffiti
'Created by Frank Verdeja
'11 November, 2003

'This project gives the user an introduction to the different types of Graffiti Art.

'This form is the opening form and main menu, which gives the user the option to begin
'the lesson, to learn terms from a glossary, look at a list of my favorite graffiti
'artitsts, or quit the project.

Private Sub Form_Load()
strPath = "n:\CS130\VBExamples\PictureForm\"

End Sub

Private Sub cmdAll_Click()
Form3.Show
Form1.Hide

End Sub

Private Sub cmdGlossary_Click()
Form12.Show
Form1.Hide

End Sub

Private Sub cmdNames_Click()
Form2.Show
Form1.Hide

End Sub

Private Sub cmdQuit_Click()
End
End Sub

