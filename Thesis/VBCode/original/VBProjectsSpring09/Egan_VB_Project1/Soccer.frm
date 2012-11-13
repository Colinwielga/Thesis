VERSION 5.00
Begin VB.Form frmSoccer 
   BackColor       =   &H00000000&
   Caption         =   "Soccer"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2085
      Left            =   1920
      Picture         =   "Soccer.frx":0000
      ScaleHeight     =   2025
      ScaleWidth      =   2025
      TabIndex        =   9
      Top             =   840
      Width           =   2085
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Homepage"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdMLS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Major League Soccer"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdLigue1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ligue 1"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdBundesliga 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bundesliga"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdLaLiga 
      BackColor       =   &H00FFFFFF&
      Caption         =   "La Liga"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSerieA 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Serie A"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdBPL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Barclays Premier League"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblWhatGame 
      BackColor       =   &H00000000&
      Caption         =   "What league would you like to bet on?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmSoccer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sports Betting Project
'frmSoccer
'Written by: Sean Egan
'Written on: 3/21/09
'This form serves as the homepage for all of the soccer betting. It
' provides links to all of the soccer leagues the user is able
' to bet on.

Private Sub cmdBPL_Click()
    'Loads BPL form
    frmBPL.Show
    'Hides Soccer homepage
    frmSoccer.Hide
End Sub

Private Sub cmdBundesliga_Click()
    'Loads Bundesliga form
    frmBundesliga.Show
    'Hides Soccer homepage
    frmSoccer.Hide
End Sub

Private Sub cmdExit_Click()
    'Closes the program
    End
End Sub

Private Sub cmdLaLiga_Click()
    'Loads La Liga form
    frmLaLiga.Show
    'Hides Soccer homepage
    frmSoccer.Hide
End Sub

Private Sub cmdLigue1_Click()
    'Loads Ligue 1 form
    frmLigue1.Show
    'Hides Soccer homepage
    frmSoccer.Hide
End Sub

Private Sub cmdMLS_Click()
    'Loads MLS form
    frmMLS.Show
    'Hides Soccer homepage
    frmSoccer.Hide
End Sub

Private Sub cmdReturn_Click()
    'Hides Soccer homepage
    frmSoccer.Hide
    'Loads Homepage
    frmHomepage.Show
End Sub

Private Sub cmdSerieA_Click()
    'Loads Serie A form
    frmSerieA.Show
    'Hides Soccer homepage
    frmSoccer.Hide
End Sub
