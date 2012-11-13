VERSION 5.00
Begin VB.Form damage 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   660
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   10875
   Begin VB.CommandButton Command7 
      Caption         =   "Battle Again"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8640
      TabIndex        =   6
      Top             =   8040
      Width           =   1815
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      Height          =   7695
      Left            =   1680
      ScaleHeight     =   7635
      ScaleWidth      =   8835
      TabIndex        =   5
      Top             =   120
      Width           =   8895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Special"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Grab"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Punch"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kick"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back To Characters"
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   8040
      Width           =   1815
   End
End
Attribute VB_Name = "damage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Super Smash Bros.
'Opening Form
'Ryan Poster and Erik Skoe
'March 26th
'The object is to have the user become familiar with how powerful the moves are in the game.
Option Explicit
Dim total As Single
Dim defeated As Boolean


Private Sub Command1_Click()
characters.Show 'To show the character form and hide the damage form.
damage.Hide

End Sub

Private Sub Command2_Click()
defeated = False
total = total + 10
Select Case total       'To add damage to the "opponent" by using different moves.
    Case Is >= 100
        MsgBox ("Your opponent is defeated.")
        defeated = True
    Case Is >= 80
        MsgBox ("Your opponent is unconscious and has a total damage of " & total & "%.")
    Case Is >= 60
        MsgBox ("Your opponent is hurt and has a total damage of " & total & "%.")
    Case Is >= 40
        MsgBox ("Your opponent is weakened and has a total damage of " & total & "%.")
    Case Is >= 20
        MsgBox ("Your opponent is silent and has a total damage of " & total & "%.")
    Case Else
        MsgBox ("Your opponent is laughing at you and has a total damage of " & total & "%.")
End Select
If defeated Then
picResults2.Picture = LoadPicture("victory.jpg")
End If
End Sub

Private Sub Command3_Click()
defeated = False
total = total + 7
Select Case total       'To add damage to the "opponent" by using different moves.
    Case Is >= 100
        MsgBox ("Your opponent is defeated.")
        defeated = True
    Case Is >= 80
        MsgBox ("Your opponent is unconscious and has a total damage of " & total & "%.")
    Case Is >= 60
        MsgBox ("Your opponent is hurt and has a total damage of " & total & "%.")
    Case Is >= 40
        MsgBox ("Your opponent is weakened and has a total damage of " & total & "%.")
    Case Is >= 20
        MsgBox ("Your opponent is silent and has a total damage of " & total & "%.")
    Case Else
        MsgBox ("Your opponent is laughing at you and has a total damage of " & total & "%.")
End Select
If defeated Then
picResults2.Picture = LoadPicture("victory.jpg")
End If
End Sub

Private Sub Command4_Click()
defeated = False
total = total + 13
Select Case total       'To add damage to the "opponent" by using different moves.
    Case Is >= 100
        MsgBox ("Your opponent is defeated.")
        defeated = True
    Case Is >= 80
        MsgBox ("Your opponent is unconscious and has a total damage of " & total & "%.")
    Case Is >= 60
        MsgBox ("Your opponent is hurt and has a total damage of " & total & "%.")
    Case Is >= 40
        MsgBox ("Your opponent is weakened and has a total damage of " & total & "%.")
    Case Is >= 20
        MsgBox ("Your opponent is silent and has a total damage of " & total & "%.")
    Case Else
        MsgBox ("Your opponent is laughing at you and has a total damage of " & total & "%.")
End Select
If defeated Then
picResults2.Picture = LoadPicture("victory.jpg")
End If
End Sub

Private Sub Command5_Click()
defeated = False
total = total + 20
Select Case total       'To add damage to the "opponent" by using different moves.
    Case Is >= 100
        MsgBox ("Your opponent is defeated.")
        defeated = True
    Case Is >= 80
        MsgBox ("Your opponent is unconscious and has a total damage of " & total & "%.")
    Case Is >= 60
        MsgBox ("Your opponent is hurt and has a total damage of " & total & "%.")
    Case Is >= 40
        MsgBox ("Your opponent is weakened and has a total damage of " & total & "%.")
    Case Is >= 20
        MsgBox ("Your opponent is silent and has a total damage of " & total & "%.")
    Case Else
        MsgBox ("Your opponent is laughing at you and has a total damage of " & total & "%.")
End Select
If defeated Then
picResults2.Picture = LoadPicture("victory.jpg")
End If
End Sub

Private Sub Command6_Click()
End 'To end the programm.
End Sub

Private Sub Command7_Click()    'This command is to clear the last battle and start again.
If total = 0 Then
MsgBox ("You have not entered a battle yet.")
Else: total = 0
picResults2.Cls
End If
End Sub

Private Sub Form_Load() ' This will load a background picture on the damage form.
picResults2.Picture = LoadPicture(App.Path & "\" & Pics(Char))
End Sub
