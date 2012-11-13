VERSION 5.00
Begin VB.Form frmMainmenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main Menu"
   ClientHeight    =   6495
   ClientLeft      =   3615
   ClientTop       =   2850
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7875
   Begin VB.CommandButton cmdhelp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdhighscores 
      BackColor       =   &H00FFFFFF&
      Caption         =   "High Scores"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCredits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdoneplayer 
      BackColor       =   &H000000C0&
      Caption         =   "Speed Type - ER"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   720
      Picture         =   "frmMainmenu.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   6315
      TabIndex        =   5
      Top             =   480
      Width           =   6375
   End
End
Attribute VB_Name = "frmMainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CSCI 130 Games to play in class (VBProject.vbp)
'Main Menu(frmMainmenu.frm)
'Done by Allen Ng
'30 October 2005
'This form is to provide a friendly interface for the user
'to choose what he/she wants to do.

Private Sub cmdCredits_Click()
    MsgBox "All work was done by me, Allen Ng.  I recieved help from Billy Jimenez, Imad Rahal, Charles Mccarron, and RJ Notaro.  Thanks for the help.", , "Credits"
End Sub

Private Sub cmdExit_Click()
    Dim A As String
    A = InputBox("Please don't quit!  If you want to quit, type Yes. If you don't want to quit type No.", "Are you sure you want to quit?") 'Quit or not to quit.
    If A = "Yes" Or A = "yes" Or A = "yeS" Or A = "YES" Or A = "YeS" Or A = "yEs" Or A = "yES" Or A = "YEs" Then
        End
    End If
End Sub

Private Sub cmdhelp_Click()
    MsgBox "Contact akng@csbsju.edu for help", , "Help"
End Sub

Private Sub cmdhighscores_Click()
    frmMainmenu.Visible = False
    frmhighscores.Visible = True
End Sub

Private Sub cmdoneplayer_Click()
    frmMainmenu.Visible = False
    frm1player.Visible = True
End Sub

