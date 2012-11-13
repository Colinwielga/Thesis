VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   5880
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   360
      Width           =   2295
   End
   Begin VB.PictureBox Picture3 
      Height          =   855
      Left            =   240
      Picture         =   "Form1.frx":493A
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   2040
      Picture         =   "Form1.frx":4F80
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   3840
      Picture         =   "Form1.frx":7791
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00800000&
      Caption         =   "Quit"
      Height          =   855
      Left            =   7200
      TabIndex        =   4
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "Other"
      Height          =   1575
      Left            =   5640
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdSCCC 
      Caption         =   "Saint Cloud   Country Club"
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdWapicada 
      BackColor       =   &H0080FFFF&
      Caption         =   "Wapicada"
      Height          =   1575
      Left            =   2040
      TabIndex        =   1
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdAngushire 
      Caption         =   "Angushire"
      Height          =   1575
      Left            =   3840
      Picture         =   "Form1.frx":94F9
      TabIndex        =   0
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "By Meghan Ellenbecker"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   "After you play a round of golf, you can enter your score into this program, and your handicap will be calculated.  "
      Height          =   735
      Left            =   960
      TabIndex        =   10
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   $"Form1.frx":B261
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1(Golf Project.vbp)
'Form1(Form1.frm)
'Meghan Ellenbecker
'March 13, 2004
'This project determines a user's handicap when they input their score, where they played (course name), and the tee boxes in which they hit from
'All of the command buttons on this form let the user choose which form/screen they want to go to


Private Sub cmdAngushire_Click()
Angushire.Show
Form1.Hide

End Sub

Private Sub cmdOther_Click()
Other.Show
Form1.Hide

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSCCC_Click()
SCCC.Show
Form1.Hide
End Sub

Private Sub cmdWapicada_Click()
Wapicada.Show
Form1.Hide
End Sub

