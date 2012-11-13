VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00FF0000&
   Caption         =   "Introduction "
   ClientHeight    =   7500
   ClientLeft      =   2520
   ClientTop       =   1920
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   9735
   Begin VB.PictureBox pic2 
      BackColor       =   &H008080FF&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H000000FF&
         Caption         =   "About the Makers"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton cmdTickets 
         BackColor       =   &H0000C000&
         Caption         =   "Buy Your Tickets Now!"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CommandButton cmdTop 
         BackColor       =   &H000080FF&
         Caption         =   "Top 10 Disney Animated Movies Of  All Time"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CommandButton cmdGiftShop 
         BackColor       =   &H00FF0000&
         Caption         =   "Gift Shop "
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00800080&
         Caption         =   "Quit "
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6240
         Width           =   2295
      End
      Begin VB.CommandButton cmdTrivia 
         BackColor       =   &H0000FFFF&
         Caption         =   "Start Trivia Game  Now "
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.PictureBox picPage1 
      Height          =   4455
      Left            =   3600
      Picture         =   "DisneyLand Project!.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label lblPage1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "...where all your dreams come true"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      TabIndex        =   2
      Top             =   6360
      Width           =   5775
   End
   Begin VB.Label lblDisneyLandTrivia 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Welcome to Disney Land Trivia"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1140
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
frmIntro.Hide
frmAbout.Show
End Sub

'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/27/06
'Objective: The objective of this form gives users a place to start the program from.  We decided to do
'our VB project with a Disney theme for a few reasons.  As we grow older as college students, we often
'become disconnected with our childhood hobbies, pastimes, and memories.  Our hope in creating this program
'was to bring back some of those childhood memories and allow busy college students to reflect on the joy and simplicity of our early lives.'
Private Sub cmdGiftShop_Click()
frmIntro.Hide   'Allow user to Go from Intro  to Gift Shop form
frmTrivia.Hide
frmGiftShop.Show
frmTop.Hide
frmTickets.Hide
End Sub

Private Sub cmdQuit_Click()
End             'Quit Program
End Sub

Private Sub cmdTickets_Click()
frmIntro.Hide    'Allow user to go from Intro to Tickets form
frmTrivia.Hide
frmGiftShop.Hide
frmTop.Hide
frmTickets.Show

End Sub

Private Sub cmdTop_Click()
frmIntro.Hide           'Allow user to go from Intro to Top form
frmTrivia.Hide
frmGiftShop.Hide
frmTop.Hide
frmTop.Show
frmTickets.Hide

End Sub

Private Sub cmdTrivia_Click()
frmIntro.Hide              'Allow user to go from Intro to Trivia form
frmTrivia.Show
frmGiftShop.Hide
frmTop.Hide
frmTickets.Hide


End Sub



