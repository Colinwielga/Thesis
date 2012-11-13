VERSION 5.00
Begin VB.Form frmIntro 
   Caption         =   "Fun Computer Games to Play During Computer Science 130"
   ClientHeight    =   6495
   ClientLeft      =   3780
   ClientTop       =   2850
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7650
   Begin VB.PictureBox picintro 
      Height          =   7000
      Left            =   120
      Picture         =   "frmIntro.frx":0000
      ScaleHeight     =   6945
      ScaleMode       =   0  'User
      ScaleWidth      =   10030.26
      TabIndex        =   0
      Top             =   120
      Width           =   10000
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Click on the screen to begin."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSCI 130 Games to play in Class(VBProject.vbp)
'Intro Form(frmIntro.frm)
'Done by Allen Ng
'30 October 2005
'This form is used to introduce the user to the program.
'The purpose of this project is to entertain the user in CSCI 130 class.
Private Sub Form_Load()
    frmIntro.Visible = True
    frmMainmenu.Visible = False
    frm1player.Visible = False
    frmhighscores.Visible = False
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub picintro_Click()
    frmIntro.Visible = False
    frmMainmenu.Visible = True
    frm1player.Visible = False
End Sub
