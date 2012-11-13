VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H000000FF&
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrep 
      Caption         =   "Prepare"
      Height          =   1095
      Left            =   720
      TabIndex        =   4
      Top             =   3960
      Width           =   1815
   End
   Begin VB.PictureBox picDmg 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   4800
      ScaleHeight     =   3555
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton cmdDefend 
      Caption         =   "Defend"
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton CmdFight 
      Caption         =   "Fight"
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.PictureBox picMonster 
      BackColor       =   &H000000FF&
      Height          =   1935
      Left            =   4680
      ScaleHeight     =   1875
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SP As Integer, HP As Integer, EHP As Integer
Dim DMG(1 To 20) As Integer, CTR As Integer
Private Sub cmdDefend_Click()
'To show a defensive action'
picDMG.Print "You take a defensive position."
picDMG.Print "The Rat swings and misses."
End Sub

Private Sub cmdFight_Click()
'I use an array for damage and put conditionals for victory and loss'
CTR = CTR + 1
EHP = EHP - DMG(CTR)
If DMG(CTR) >= 200 Then
    picDMG.Print "CRITICAL HIT!"
ElseIf DMG(CTR) < 1 Then
    picDMG.Print "You Miss."
ElseIf DMG(CTR) = 1 - 199 Then
    picDMG.Print "You Hit the Rat for"; DMG(CTR); " Damage"
End If
If EHP <= 0 Then
        MsgBox "You Killed the Rat, you get 50 treasure."
        Treasure = Treasure + 50
        MsgBox "You Gained a level."
        Level = Level + 1
        Form4.Hide
        Form5.Show
ElseIf EHP > 0 Then
        picDMG.Print "The Rat does 2 damage to you."
        HP = HP - 2
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
End If
If CTR = 20 Then CTR = 1
End Sub

Private Sub cmdPrep_Click()
'Used to sort out the best damage and let the player use that for their next attack'
Dim Pos As Integer, Pass As Integer
Dim TDMG As Integer
CTR = 20
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If DMG(Pos) < DMG(Pos + 1) Then
            TDMG = DMG(Pos)
            DMG(Pos) = DMG(Pos + 1)
            DMG(Pos + 1) = TDMG
        End If
    Next Pos
Next Pass
CTR = 0
picDMG.Print "You prepare your next attack. It looks like it's going to hurt."
picDMG.Print "The Rat does 1 damage to you."

HP = HP - 1
picDMG.Print ; HP; ": Hitpoints left"
If HP <= 0 Then
    MsgBox "You Die. Game over. Please Try again."
    End
End If
End Sub

Private Sub Form_Load()
'Start of the first fight, what I need to start it.'
'I need to load a damage array, and various stats so I can calculate victory'
CTR = 0
Open Path & "magedmg.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, DMG(CTR)
    Loop
Close
HP = 50
SP = 5
EHP = 20
CTR = 0
picMonster.Picture = LoadPicture("Rat.jpg")
picDMG.Print "A Large Rat appears and begins to attack you."
End Sub

Private Sub Picture1_Click()

End Sub
