VERSION 5.00
Begin VB.Form frmMainmenu 
   BackColor       =   &H00400000&
   Caption         =   "Main Menu"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   Picture         =   "frmMainmenu.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgotoform5 
      Caption         =   "Fall SJU Rugby Players"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit Program"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      TabIndex        =   3
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdgotoform4 
      Caption         =   "History of Rugby"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdgotoform3 
      Caption         =   "Fall SJU Rugby Scores"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2160
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdgotoform2 
      Caption         =   "Position Explanation"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   6120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : RugbyVBProject (Rugby.vbp)
'Form Name : frmMainmenu(frmMainmenu.frm)
'Author: Alex Forrest
'Date : Monday October 31, 2005
'Purpose of the Project: Upon user interaction, this program is directed to help the
    'user gain a basic knowledge for the positions of Rugby and the history of the game
    'itself.  It also incorporates Saint John's Rugby by allowing the user to view the
    'scores of the SJU rugby team this year along with the players on the team.
'Purpose of the form:  It is the foundation of the project.  It allows the user to
    'see the general aspect of the program along with what else will be included in the
    'program.  It includes five command buttons, four of which take the user to other
    'forms, and one button that allows the user to quit the program.

Private Sub cmdgotoform2_Click()
    frmMainmenu.Hide 'hides the main menu
    frmpositions.Show 'takes the user to the position form to explain the positions to the user.
End Sub

Private Sub cmdgotoform3_Click()
    frmMainmenu.Hide
    frmSJUscores.Show 'takes the user to the scores form to show the user the scores of the SJU rugby team.
End Sub

Private Sub cmdgotoform4_Click()
    frmMainmenu.Hide
    frmhistory.Show 'takes the user to the history form to explain the history of rugby.
End Sub

Private Sub cmdgotoform5_Click()
    frmMainmenu.Hide
    frmplayers.Show 'takes the user to the player form and shows the players on the team
End Sub

Private Sub cmdquit_Click()
    End 'allows the user to quit the program.
End Sub

