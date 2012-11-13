VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FF0000&
   Caption         =   "Form11"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form11"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDungeon 
      BackColor       =   &H00FF0000&
      Height          =   3375
      Left            =   4920
      ScaleHeight     =   3315
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton cmdNorth 
      Caption         =   "North"
      Height          =   975
      Left            =   1800
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdSleep 
      Caption         =   "Sleep"
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   $"Form11.frx":0000
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HP As Integer
Private Sub cmdNorth_Click()
'Form loader to show movement'
Form11.Hide
Form12.Show
End Sub

Private Sub cmdSleep_Click()
'this function is to show rest, but if they use it too long they die'
Dim Sleep As Integer

Sleep = Sleep + 1
If Sleep < 4 Then
    HP = HP + 5
    MsgBox ("You gained 5 hit points bringing your health to:" & HP&)
ElseIf Sleep >= 4 Then
    MsgBox ("You slept too long, and were killed in your sleep.")
    End
End If
End Sub

Private Sub Form_Load()
'used for variables and pictures.'
HP = 45
picDungeon.Picture = LoadPicture("Dungeon7d.jpg")
MsgBox ("As you take care of the zombie, rocks in the room fall preventing you from going back the way you came.")
End Sub
