VERSION 5.00
Begin VB.Form Intro1 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Quit1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quit"
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Continue1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click Here to Continue..."
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"VBProject4.frx":0000
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   5055
   End
End
Attribute VB_Name = "Intro1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Continue1_Click()

' Which not so nice person would you like to beat up today? "Megan'sVBProject.vbp"
'                       Intro1 (VBProject4.frm)
'                       Megan Kelly 11/03/03
' Project Purpose:  The purpose of this project is to provide some small form of entertainment to people who just don't like evil.
'       The scores of the 5 opponents were created by finding the scores that they would get if they were to take this test.  They were all given the age of 42.
' Purpose:  The purpose of this form is to provide an introduction to the program, and to read in the names & scores of all opponents.


Intro1.Visible = False
Open "M:/cs130/MeganKellyProject/" & "namefactor.txt" For Input As #1
    For k = 1 To 5
        Input #1, opponentname(k), opponentfactor(k)
    Next k
Close #1
mansonwins.Visible = True
SelectOpponent2.Visible = True
End Sub

Private Sub Quit1_Click()
End
End Sub