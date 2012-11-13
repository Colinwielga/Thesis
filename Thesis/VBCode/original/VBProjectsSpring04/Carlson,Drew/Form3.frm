VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF0000&
   Caption         =   "MIAC Champs Picture"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form3"
   ScaleHeight     =   5655
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Click to end SJU hockey stats program"
      Height          =   855
      Left            =   5640
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Click to return to main page of SJU hockey stats program"
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click to return to Formal team picture"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   0
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Created by: Drew Carlson"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form3 (Form3.frm)
'This form exists to show another picture of the team.



Private Sub Command1_Click()
'This button returns user to the first picture of the team.
Form2.Show
Form3.Hide

End Sub

Private Sub Command2_Click()
'This button returns the user back to the main page.
Form1.Show
Form3.Hide
End Sub

Private Sub Command3_Click()
'This button quits the program.
MsgBox ("Thank you for using the SJU Hockey Statistic Program.")
End
End Sub
