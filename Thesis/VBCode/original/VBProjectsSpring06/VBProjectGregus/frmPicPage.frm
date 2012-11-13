VERSION 5.00
Begin VB.Form frmPicPage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Picture Page"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExplain 
      Caption         =   "What's going on here?"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   6840
      Width           =   3135
   End
   Begin VB.PictureBox picBoxTexas 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   1320
      Picture         =   "frmPicPage.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   2895
      TabIndex        =   7
      Top             =   7800
      Width           =   2895
   End
   Begin VB.PictureBox picBoxAfter 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4680
      Picture         =   "frmPicPage.frx":1CDD
      ScaleHeight     =   2055
      ScaleWidth      =   4095
      TabIndex        =   6
      Top             =   7800
      Width           =   4095
   End
   Begin VB.CommandButton cmbBackHome 
      BackColor       =   &H8000000B&
      Caption         =   "Back to Home"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9960
      MaskColor       =   &H00400040&
      TabIndex        =   5
      Top             =   8880
      Width           =   1695
   End
   Begin VB.PictureBox picBoxMud4 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   6960
      Picture         =   "frmPicPage.frx":4274
      ScaleHeight     =   3615
      ScaleWidth      =   4575
      TabIndex        =   4
      Top             =   3720
      Width           =   4575
   End
   Begin VB.PictureBox picBoxMud3 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      Picture         =   "frmPicPage.frx":5AC86
      ScaleHeight     =   2895
      ScaleWidth      =   6855
      TabIndex        =   3
      Top             =   3720
      Width           =   6855
   End
   Begin VB.PictureBox picBoxMud2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   9120
      Picture         =   "frmPicPage.frx":5E8F4
      ScaleHeight     =   2055
      ScaleWidth      =   6855
      TabIndex        =   2
      Top             =   1440
      Width           =   6855
   End
   Begin VB.PictureBox picBoxMud1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      Picture         =   "frmPicPage.frx":61AE8
      ScaleHeight     =   2295
      ScaleWidth      =   9015
      TabIndex        =   1
      Top             =   1200
      Width           =   9015
   End
   Begin VB.PictureBox frmPicPage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      Picture         =   "frmPicPage.frx":66A73
      ScaleHeight     =   975
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label lblCredit 
      BackColor       =   &H00000000&
      Caption         =   "Project by: Dan Gregus"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblDescription3 
      BackColor       =   &H00000000&
      Caption         =   "It's good to be home, even in the middle of February."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4800
      TabIndex        =   9
      Top             =   9840
      Width           =   3735
   End
   Begin VB.Label LblDescription2 
      BackColor       =   &H00000000&
      Caption         =   "Tournamnets in Texas where it is 78 degrees and sunny is always  nice but..."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1320
      TabIndex        =   8
      Top             =   9960
      Width           =   3135
   End
End
Attribute VB_Name = "frmPicPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmPicPage (frmPicPage.frm)
'Dan Gregus
'3/21/06
'Objective: To showcase the lighter side of the SJU lacross team.  If given more time I would like to expand this page to include more pictures, more pages, and pictures that provide sounds when clicked.

Private Sub cmbBackHome_Click()
    frmPicPage.Visible = False
    frmSJULacrosse.Visible = True
End Sub


Private Sub cmdExplain_Click()
MsgBox "John Pinball Carlson helps lighten the mood after a rainy game day with his now infamous mudslide.", , "Just having Fun"
End Sub

