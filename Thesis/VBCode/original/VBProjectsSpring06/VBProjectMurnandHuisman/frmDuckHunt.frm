VERSION 5.00
Begin VB.Form frmduckhunt 
   Caption         =   "DuckHunt"
   ClientHeight    =   4590
   ClientLeft      =   480
   ClientTop       =   1305
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   9390
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   0
      Picture         =   "frmDuckHunt.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   9315
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton cmdStart 
         Caption         =   "Push my buttons"
         Height          =   1695
         Left            =   2640
         TabIndex        =   1
         Top             =   2880
         Width           =   3615
      End
      Begin VB.CommandButton cmdQuick 
         Caption         =   "Click Here Quickly for Results"
         Height          =   1215
         Left            =   3120
         TabIndex        =   11
         Top             =   2880
         Width           =   3135
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdClick7 
         Height          =   975
         Left            =   5400
         Picture         =   "frmDuckHunt.frx":A8F12
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdClick6 
         Height          =   975
         Left            =   6000
         Picture         =   "frmDuckHunt.frx":A94C6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdClick5 
         Height          =   975
         Left            =   3480
         Picture         =   "frmDuckHunt.frx":A9A7A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdClick4 
         Height          =   975
         Left            =   1680
         Picture         =   "frmDuckHunt.frx":AA02E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdClick3 
         Height          =   975
         Left            =   1680
         Picture         =   "frmDuckHunt.frx":AA5E2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdclick2 
         Height          =   975
         Left            =   7920
         Picture         =   "frmDuckHunt.frx":AAB96
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdClick1 
         Height          =   975
         Left            =   7560
         Picture         =   "frmDuckHunt.frx":AB14A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lblnames 
         Caption         =   "Designed by: CJ and Murn"
         Height          =   255
         Left            =   6360
         TabIndex        =   12
         Top             =   4080
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmduckhunt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name : Trapshooting(PCJ Jim Project.vbp)
'Form Name : frmCJandJim(frmCJandJim.frm)
'Author: James Murn & Chelsey Jo Huisman
'Date : Wednesday March 22, 2006
'Purpose of this form:  This form allows the users to
                       'do absolutely nothing
                       'Murn seems to think that he is creative and designed this worthless form
                       'it is created purly for the users enjoyment
                       'please enjoy
                       'to skip to the next form
                       'to exit the program
                       'to start the simple game
Private Sub cmdclick2_Click()
    cmdClick1.Visible = False
    cmdclick2.Visible = False 'used to show the next button in the game,while hidding all the others
    cmdClick3.Visible = True
    cmdClick4.Visible = False
    cmdClick5.Visible = False
    cmdClick6.Visible = False
    cmdClick7.Visible = False
    cmdQuick.Visible = False
End Sub

Private Sub cmdClick3_Click()
    cmdClick1.Visible = False
    cmdclick2.Visible = False
    cmdClick3.Visible = False
    cmdClick4.Visible = True 'used to show the next button in the game,while hidding all the others
    cmdClick5.Visible = False
    cmdClick6.Visible = False
    cmdClick7.Visible = False
    cmdQuick.Visible = False
End Sub

Private Sub cmdClick4_Click()
    cmdClick1.Visible = False
    cmdclick2.Visible = False
    cmdClick3.Visible = False
    cmdClick4.Visible = False
    cmdClick5.Visible = True 'used to show the next button in the game,while hidding all the others
    cmdClick6.Visible = False
    cmdClick7.Visible = False
    cmdQuick.Visible = False
End Sub

Private Sub cmdClick5_Click()
    cmdClick1.Visible = False
    cmdclick2.Visible = False
    cmdClick3.Visible = False
    cmdClick4.Visible = False
    cmdClick5.Visible = False
    cmdClick6.Visible = True 'used to show the next button in the game,while hidding all the others
    cmdClick7.Visible = False
    cmdQuick.Visible = False
End Sub

Private Sub cmdClick6_Click()
    cmdClick1.Visible = False
    cmdclick2.Visible = False
    cmdClick3.Visible = False
    cmdClick4.Visible = False
    cmdClick5.Visible = False
    cmdClick6.Visible = False
    cmdClick7.Visible = True 'used to show the next button in the game,while hidding all the others
    cmdQuick.Visible = False
End Sub

Private Sub cmdClick7_Click()
    cmdClick1.Visible = False
    cmdclick2.Visible = False
    cmdClick3.Visible = False
    cmdClick4.Visible = False
    cmdClick5.Visible = False
    cmdClick6.Visible = False
    cmdClick7.Visible = False
    cmdQuick.Visible = True 'used to show the next button in the game,while hidding all the others
End Sub

Private Sub cmdExit_Click()
    End 'allows the user to exit the program
End Sub

Private Sub cmdNext_Click()
    frmCJandJim.Show
    frmduckhunt.Hide 'used to go back to the first form
End Sub

Private Sub cmdQuick_Click()
    cmdClick1.Visible = False
    cmdclick2.Visible = False
    cmdClick3.Visible = False
    cmdClick4.Visible = False
    cmdClick5.Visible = False 'used to display the users result in a msg box
    cmdClick6.Visible = False
    cmdClick7.Visible = False
    MsgBox "You shot 'em all!!", , "Your Virtual DuckHunt" 'these are some confidence boosters just in case you shot bad at the club today
    MsgBox "Yes you did!!", , "Your Personal Virtual DuckHunt"
    MsgBox "Click the Exit box NOW!!", , "It's time to END the game"
End Sub

Private Sub cmdStart_Click()
    cmdClick1.Visible = True 'used to start the game and show the next button in the game,while hidding all the others
    cmdclick2.Visible = False
    cmdClick3.Visible = False
    cmdClick4.Visible = False
    cmdClick5.Visible = False
    cmdClick6.Visible = False
    cmdClick7.Visible = False
    cmdQuick.Visible = False
    cmdStart.Visible = False
End Sub

Private Sub cmdClick1_Click()
    cmdClick1.Visible = False
    cmdclick2.Visible = True 'used to show the next button in the game,while hidding all the others
    cmdClick3.Visible = False
    cmdClick4.Visible = False
    cmdClick5.Visible = False
    cmdClick6.Visible = False
    cmdClick7.Visible = False
    cmdQuick.Visible = False

End Sub
