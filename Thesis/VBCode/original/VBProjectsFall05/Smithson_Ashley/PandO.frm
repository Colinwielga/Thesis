VERSION 5.00
Begin VB.Form PennOrient 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PennOrient"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What is the best kept secret of the P and O?"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdhow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How Many Bedrooms and Bathrooms are in the P and O?"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdwhat2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What Does P and O Stand For?"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   240
      Picture         =   "PandO.frx":0000
      ScaleHeight     =   7035
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   1320
      Width           =   5535
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to Home Page"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton cmdwho 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Who lives there?"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdwhere 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Where is it located?"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdwhat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What is the  P and O?"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Ashley 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label PANDO 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P AND O HOTEL        (quick facts)"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "PennOrient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Australia
'Form Name: PennOrient
'Author: Ashley Smithson
'Date: October 31, 2005
'Purpose of form: show some interesting information about the P & O hotel.
Option Explicit
Private Sub cmdback_Click()
PennOrient.Hide 'brings user back to the home page
FinalProject2.Show
End Sub

Private Sub cmdbest_Click()
MsgBox "The best kept secret in the P & O is the Movie Theatre in the basement!", , "Secret"
'displays the answer to the question displayed on the button
End Sub

Private Sub cmdhow_Click()
MsgBox "The P & O has 30 bedrooms and 4 bathrooms, with 6 showers and 7 toilets, for 40 people.", , "Bedrooms and Bathrooms"
End Sub

Private Sub cmdwhat_Click()
MsgBox "The P & O is a restored hotel, converted to a dorming facility for study abroad students.", , "P and O"
End Sub

Private Sub cmdwhat2_Click()
MsgBox "Pennisula and Orient", , "P and O"
End Sub

Private Sub cmdwhere_Click()
MsgBox "The P & O is located on the corner of Moat and High Street in Fremantle, Western Australia.", , "Freo"
End Sub

Private Sub cmdwho_Click()
MsgBox "The P & O houses study abroad students, in the spring of 2005 it housed the people you will meet.", , "Friends"
End Sub
