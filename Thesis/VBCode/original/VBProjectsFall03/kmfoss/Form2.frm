VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   Caption         =   "Form2"
   ClientHeight    =   8115
   ClientLeft      =   285
   ClientTop       =   660
   ClientWidth     =   11760
   LinkTopic       =   "Form2"
   ScaleHeight     =   8115
   ScaleWidth      =   11760
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   3120
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   5355
      TabIndex        =   7
      Top             =   1200
      Width           =   5415
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9720
      TabIndex        =   6
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdSkills 
      Caption         =   "Skills and Benefits"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   5
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCollateral 
      Caption         =   "Collateral Assignments"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5880
      TabIndex        =   4
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdInteraction 
      Caption         =   "Resident Interaction"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdDuty 
      Caption         =   "Duty"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      TabIndex        =   2
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdProgramming 
      Caption         =   "Programming"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      TabIndex        =   1
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form2.frx":9BFC
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCollateral_Click()
Form2.Hide
Form6.Show
End Sub

Private Sub cmdDuty_Click()
Form2.Hide
Form5.Show
End Sub

Private Sub cmdINteraction_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub cmdProgramming_Click()
Form2.Hide
Form4.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRequirements_Click()
Form2.Hide
Form8.Show
End Sub

Private Sub cmdSkills_Click()
Form2.Hide
Form7.Show
End Sub
