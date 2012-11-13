VERSION 5.00
Begin VB.Form frmhighscores 
   BackColor       =   &H00FFFFFF&
   Caption         =   "High Scores"
   ClientHeight    =   6495
   ClientLeft      =   2970
   ClientTop       =   3015
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Picture         =   "frmhighscores.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   8985
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdrace 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Speed Type-ER High Scores"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1800
   End
   Begin VB.PictureBox picoutput 
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6195
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label lblshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on the Button to view the high scores."
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "frmhighscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CSCI 130 Games to play in class (VBProject.vbp)
'High Scores(frmhighscores.frm)
'Done by Allen Ng
'30 October 2005
'This form is to provide a friendly interface for the user
'to view the high scores from the program.
Private Sub cmdback_Click()
    frmhighscores.Visible = False
    frmMainmenu.Visible = True
End Sub

Private Sub cmdrace_Click()
    picoutput.Cls
    picoutput.Print "Name", Tab(40); "Score"
    picoutput.Print "*******************************************************"
    Open App.Path & "\Hall of Fame.txt" For Input As #4
    For I = 1 To 10
        Input #4, Halloffamename(I), Halloffamescore(I)
        picoutput.Print Halloffamename(I), Tab(40); Halloffamescore(I)
        picoutput.Print
    Next I
    Close #4
End Sub

