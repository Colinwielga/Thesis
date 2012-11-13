VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF0000&
   Caption         =   "Johnnie Hockey Team Picture"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   5715
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdend 
      Caption         =   "Click to end SJU hockey stats program"
      Height          =   855
      Left            =   6240
      TabIndex        =   3
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton cmdtitle 
      Caption         =   "Click to see SJU hockey team after winning MIAC title"
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Click to return to main page of SJU hockey stats program"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      Picture         =   "Form2.frx":3FF8A
      ScaleHeight     =   4395
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Created by: Drew Carlson"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form2 (Form2.frm)
'This form exists to show the user the picture of the team.
Private Sub cmdend_Click()
'This button quits the program.
MsgBox ("Thank you for using the SJU Hockey Statistic Program.")
End
End Sub

Private Sub cmdreturn_Click()
'This button brings the user back to the starting form
Form1.Show
Form2.Hide
End Sub

Private Sub cmdtitle_Click()
'This button switches forms to display different picture.
Form3.Show
Form2.Hide
End Sub

Private Sub Command1_Click()
' This button brings user back to the main page.
Form1.Show
Form2.Hide

End Sub

