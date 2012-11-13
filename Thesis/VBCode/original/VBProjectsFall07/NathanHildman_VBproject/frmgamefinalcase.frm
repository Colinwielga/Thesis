VERSION 5.00
Begin VB.Form frmgamefinalcase 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdyourcase 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click to see what your case holds"
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdfinalcase 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click to what the final case holds"
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.PictureBox picwinning 
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
      Height          =   1695
      Left            =   3120
      ScaleHeight     =   1635
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
End
Attribute VB_Name = "frmgamefinalcase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdfinalcase_Click()
picwinning.Print "You have won"; casefinal;                 'displays how much money was in the case selected
cmdyourcase.Visible = True                                  'hides one button displays the two others
cmdfinalcase.Visible = False
cmdquit.Visible = True


End Sub

Private Sub cmdquit_Click()
End                                                            'ends program
End Sub

Private Sub cmdyourcase_Click()
frmgamefinalcase.Hide                                   'goes to another program
frmplayerscase.Show

End Sub

