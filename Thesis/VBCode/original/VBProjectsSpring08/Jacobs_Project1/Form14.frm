VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H0000FF00&
   Caption         =   "Form14"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form14"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEast 
      Caption         =   "East"
      Height          =   975
      Left            =   2760
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdWest 
      Caption         =   "West"
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPotions 
      Caption         =   "Drink a potion"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   $"Form14.frx":0000
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEast_Click()
'This  moves them to the last fight of the game'
MsgBox ("You better be ready, because here comes the fight of your life.")
Form14.Hide
Form15.Show
End Sub

Private Sub cmdPotions_Click()
Dim Pass As Integer, Pos As Integer
Dim TDMG As Integer, Potion As String
'used to add a little puzzle element to the game'
Potion = InputBox("Select a potion. (Green, Blue, Red, or Pink are your options)")

SelectCase Potion
    Case Is = Green
        MsgBox ("You die. Game over")
        End
    Case Is = Blue
        MsgBox ("You turn into a fish, and promptly die from lack of water. Game over.")
        End
    Case Is = Red
        MsgBox ("You feel enraged")
    Case Is = Pink
        MsgBox ("You start floating in the air, you hit your head on the ceiling, and pass out")
        HP = HP - 10
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
        End
        End If
    Case Is = UNSEEN
        MsgBox ("You find the hidden potion and drink it. You feel amazing.")
        HP = HP + 50
        SP = SP + 5
    Case Else
        MsgBox ("Pick an actual potion.")
    End Select
End Sub

Private Sub cmdWest_Click()
'Movement function'
Form14.Hide
Form13.Show
End Sub
