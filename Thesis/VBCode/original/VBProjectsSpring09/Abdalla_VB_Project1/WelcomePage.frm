VERSION 5.00
Begin VB.Form FrmWelcome 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   5280
      TabIndex        =   5
      Top             =   6720
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   3720
      Picture         =   "WelcomePage.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton cmdTeacher 
      Caption         =   "Proffesor"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdBuyer 
      Caption         =   "Buyer"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdSeller 
      Caption         =   "seller"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Welcome to Price your Book"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "FrmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Book exchange'
'Form name: frmWelcome'
'Author: Bibi Abdalla'
'Date: 3/24/2009'
'Objective: Wecome Users and direct them  to the approciate form'
'over all objective of project: enable students that are books and buys to communicate. Also allow professor to register courses and required books so that their students know which books to buy or sell.'


Private Sub cmdBuyer_Click()
    'hide welcome page and show Bookswap page'
    FrmWelcome.Hide
    frmBookswap.Show

End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSeller_Click()
    'hide welcome page and show the seller page'
    FrmWelcome.Hide
    frmSeller.Show

End Sub

Private Sub cmdTeacher_Click()
    'hide welcome page and show Professor page'
    FrmWelcome.Hide
    frmProf.Show

End Sub

