VERSION 5.00
Begin VB.Form frmplayerscase 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0000FFFF&
      Caption         =   "quit"
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdothercase 
      BackColor       =   &H0000FFFF&
      Caption         =   "See what was in the other case"
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdyourcase 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click to open your case"
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox picplayerscase 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3240
      ScaleHeight     =   1755
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   1680
      Width           =   5055
   End
End
Attribute VB_Name = "frmplayerscase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdquit_Click()
End                                             'ends program
End Sub

Private Sub cmdyourcase_Click()
picplayerscase.Print "You Have Won"; playerscase;    'displays the value you had in your case
cmdyourcase.Visible = False                           'hides one button and displays two others
cmdothercase.Visible = True
cmdquit.Visible = True



End Sub

Private Sub cmdothercase_Click()                        'goes to another form

frmplayerscase.Hide
frmplayerscase.Show
End Sub


