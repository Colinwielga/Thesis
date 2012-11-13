VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H000000FF&
   Caption         =   "Form10"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form10"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDMG 
      BackColor       =   &H000000FF&
      Height          =   2535
      Left            =   5160
      ScaleHeight     =   2475
      ScaleWidth      =   3435
      TabIndex        =   6
      Top             =   3600
      Width           =   3495
   End
   Begin VB.PictureBox picMonster 
      BackColor       =   &H000000FF&
      Height          =   2775
      Left            =   4440
      ScaleHeight     =   2715
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton cmdPrep 
      Caption         =   "Prepare"
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "Fire Ball"
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdbolt 
      Caption         =   "Lightning Bolt"
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdDefend 
      Caption         =   "Defend"
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "Fight"
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SP As Integer, HP As Integer, EHP As Integer
Dim DMG(1 To 20) As Integer, CTR As Integer
Private Sub cmdBolt_Click()
'This is a spell, to damage the enemy.'
If SP >= 5 Then
    picDMG.Print "The Zombie takes 50 dmg"
    EHP = EHP - 50
    SP = SP - 5
    picDMG.Print ; SP; ": spell points remaining."
    If EHP <= 0 Then
        MsgBox "You Killed the Zombie, you get 250 treasure."
        Treasure = Treasure + 250
        MsgBox "You gained a level."
        Level = Level + 1
        Form10.Hide
        Form11.Show
    ElseIf EHP > 0 Then
        picDMG.Print "The Zombie does 12 damage to you."
        HP = HP - 12
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
    End If
ElseIf SP < 5 Then
    MsgBox "You do not have enough Magic points to cast this spell."
End If
End Sub

Private Sub cmdDefend_Click()
'to show a defensive action'
picDMG.Print "You take a defensive position."
picDMG.Print "The Zombie swings and misses."
End Sub

Private Sub cmdFight_Click()
'This is how I calculate melee damage'
CTR = CTR + 1
EHP = EHP - DMG(CTR)
If DMG(CTR) >= 200 Then
    picDMG.Print "CRITICAL HIT!"
ElseIf DMG(CTR) < 1 Then
    picDMG.Print "You Miss."
ElseIf DMG(CTR) = 1 - 199 Then
    picDMG.Print "You Hit the Zombie for"; DMG(CTR); " Damage"
End If
If EHP <= 0 Then
        MsgBox "You Killed the Zombie, you get 250 treasure."
        Treasure = Treasure + 250
        Form10.Hide
        Form11.Show
ElseIf EHP > 0 Then
        picDMG.Print "The Zombie does 10 damage to you."
        HP = HP - 10
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
End If
If CTR = 20 Then CTR = 1
End Sub

Private Sub cmdFire_Click()
'This is a spell, to damage the enemy.'
If SP >= 20 Then
    picDMG.Print "The Zombie takes 200 dmg"
    EHP = EHP - 200
    SP = SP - 20
    picDMG.Print ; SP; ": spell points remaining."
    If EHP <= 0 Then
        MsgBox "You Killed the Zombie, you get 250 treasure."
        Treasure = Treasure + 250
        Form10.Hide
        Form11.Show
    ElseIf EHP > 0 Then
        picDMG.Print "The Zombie does 12 damage to you."
        HP = HP - 12
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
    End If
ElseIf SP < 20 Then
    MsgBox "You do not have enough Magic points to cast this spell."
End If

End Sub

Private Sub cmdPrep_Click()
'This is to create the next attack as the hightest possible'
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
picDMG.Print "The Zombie does 9 damage to you."

HP = HP - 9
picDMG.Print ; HP; ": Hitpoints left"

End Sub


Private Sub Form_Load()
'I use this to show the monster, and set global variables'
CTR = 0
Open Path & "theifdmg.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, DMG(CTR)
Loop
Close
HP = 70
SP = 30
EHP = 225
CTR = 6
picmonster.Picture = LoadPicture("zombie.jpg")
picDMG.Print "A zombie spots you and begins to attack."
End Sub
