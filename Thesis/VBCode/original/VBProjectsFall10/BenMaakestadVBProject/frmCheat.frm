VERSION 5.00
Begin VB.Form frmCheat 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   8760
      Top             =   1680
   End
   Begin VB.PictureBox picCheat 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   360
      ScaleHeight     =   6555
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   2550
      Left            =   7800
      Picture         =   "frmCheat.frx":0000
      Top             =   3000
      Width           =   4455
   End
End
Attribute VB_Name = "frmCheat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer2_Timer()
frmCheat.Hide   'after a few seconds this function returns the user to the puzzle form

Timer2.Enabled = False
frmPuzzle.Show


frmPuzzle.picPuzzle.Cls
For Pos = 1 To 81
        
        frmPuzzle.picPuzzle.Print Puz(Pos); "    "; Puz(Pos + 1); "    "; Puz(Pos + 2); "    "; Puz(Pos + 3); "    "; Puz(Pos + 4); "    "; Puz(Pos + 5); "    "; Puz(Pos + 6); "    "; Puz(Pos + 7); "    "; Puz(Pos + 8)
        frmPuzzle.picPuzzle.Print Tab(50)
        Pos = Pos + 8
        
Next Pos
   

End Sub
