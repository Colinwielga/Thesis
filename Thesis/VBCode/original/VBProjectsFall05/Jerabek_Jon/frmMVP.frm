VERSION 5.00
Begin VB.Form frmMVP 
   BackColor       =   &H00800000&
   Caption         =   "MVP"
   ClientHeight    =   6030
   ClientLeft      =   3540
   ClientTop       =   3075
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8535
   Visible         =   0   'False
   Begin VB.PictureBox picOutput1 
      Height          =   855
      Left            =   360
      ScaleHeight     =   795
      ScaleWidth      =   6315
      TabIndex        =   6
      Top             =   5040
      Width           =   6375
   End
   Begin VB.PictureBox picOutput 
      Height          =   855
      Left            =   5520
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdGuess 
      BackColor       =   &H8000000E&
      Caption         =   "Guess"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Height          =   855
      Left            =   5280
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox picTrophy 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   360
      Picture         =   "frmMVP.frx":0000
      ScaleHeight     =   4575
      ScaleWidth      =   3735
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton cmdMain1 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   0
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Guess the season(yy-yy) that Kevin Garnett won the NBA MVP Award"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmMVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ProjectKG
'frmMVP
'Jon Jerabek
'10-25-05 & 10-26-05
'Objective-Allows user to guess the season KG won the MVP. Also displays facts if correct.

Private Sub cmdGuess_Click()
Dim x As String
x = txtYear.Text               'User inputs season
picOutput.Cls
If x = "03-04" Then            'Compares the user input with the correct input and displays appropriate result
    MsgBox "CORRECT!!", , "Woohoo!"
    picOutput.Print "Correct!"
    picOutput1.Print "Kevin Garnett was named the 03-04 NBA MVP on May 3, 2004."
    picOutput1.Print "He is the only player from the Timberwolves to ever receive the honor."
    picOutput1.Print "He is also one of two men to ever come straight from high school and win MVP."
    Else
        MsgBox "WRONG!", , "Sorry!"
        picOutput.Print "Wrong!"
        picOutput.Print "Try again"
End If

End Sub

Private Sub cmdMain1_Click()
frmHome.Show
frmMVP.Hide
End Sub
