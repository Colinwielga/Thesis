VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H000000FF&
   Caption         =   "Form15"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form15"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSurrender 
      Caption         =   "Surrender!"
      Height          =   855
      Left            =   3240
      TabIndex        =   9
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdSneak 
      Caption         =   "Sneak by...or at least try."
      Height          =   855
      Left            =   3240
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrep 
      Caption         =   "Prepare"
      Height          =   855
      Left            =   3240
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "FIREBALL"
      Height          =   1095
      Left            =   480
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdIce 
      Caption         =   "ICE LANCE!"
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdBolt 
      Caption         =   "LIGHTNING BOLT"
      Height          =   1095
      Left            =   480
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdDefend 
      Caption         =   "Defend"
      Height          =   1095
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "Fight"
      Height          =   1335
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox picDMG 
      BackColor       =   &H000000FF&
      Height          =   3855
      Left            =   5760
      ScaleHeight     =   3795
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   4680
      Width           =   4575
   End
   Begin VB.PictureBox picmonster 
      BackColor       =   &H000000FF&
      Height          =   3855
      Left            =   3600
      ScaleHeight     =   3795
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   480
      Width           =   6495
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SP As Integer, HP As Integer, EHP As Integer
Dim DMG(1 To 20) As Integer, CTR As Integer
Private Sub cmdBolt_Click()
'spell damage'
If SP >= 5 Then
    picDMG.Print "The Dragon takes 50 dmg"
    EHP = EHP - 50
    SP = SP - 5
    picDMG.Print ; SP; ": spell points remaining."
    If EHP <= 0 Then
        MsgBox "You Killed the Dragon, you get 15000 treasure."
        Treasure = Treasure + 15000
        MsgBox "You Gained a level."
        Level = Level + 1
        Form15.Hide
        Form16.Show
    ElseIf EHP > 0 Then
        picDMG.Print "The Dragon does 40 damage to you."
        HP = HP - 40
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
'Defensive action'
picDMG.Print "You take a defensive position."
picDMG.Print "The Dragon calls you a pathetic bug, and flicks you with his claw."
HP = HP - 15
picDMG.Print ; HP; ": Hitpoints left"
If HP <= 0 Then
    MsgBox "You Die. Game over. Please Try again."
    End
End If
End Sub

Private Sub cmdFight_Click()
'Standard melee damage'
CTR = CTR + 1
EHP = EHP - DMG(CTR)
If DMG(CTR) >= 200 Then
    picDMG.Print "CRITICAL HIT!"
ElseIf DMG(CTR) < 1 Then
    picDMG.Print "You Miss."
ElseIf DMG(CTR) = 1 - 199 Then
    picDMG.Print "You Hit the Dragon for"; DMG(CTR); " Damage"
End If
If EHP <= 0 Then
        MsgBox "You Killed the Dragon, you get 15000 treasure."
        Treasure = Treasure + 15000
        MsgBox "You Gained a level."
        Level = Level + 1
        Form13.Hide
        Form14.Show
ElseIf EHP > 0 Then
        picDMG.Print "The Dragon does 50 damage to you."
        HP = HP - 20
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
End If
If CTR = 20 Then CTR = 0
End Sub

Private Sub cmdFire_Click()
'spell damage'
If SP >= 20 Then
    picDMG.Print "The Dragon takes 200 dmg"
    EHP = EHP - 200
    SP = SP - 20
    picDMG.Print ; SP; ": spell points remaining."
    If EHP <= 0 Then
        MsgBox "You Killed the Dragon, you get 15000 treasure."
        Treasure = Treasure + 15000
        MsgBox "You Gained a level."
        Level = Level + 1
        Form15.Hide
        Form16.Show
    ElseIf EHP > 0 Then
        picDMG.Print "The Dragon does 40 damage to you."
        HP = HP - 40
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

Private Sub cmdIce_Click()
'Spell for an instant kill'
If SP >= 50 Then
    picDMG.Print "The Ice Lance hits the Dragon through the heart. Slaying it instantly."
    EHP = EHP - 1000
    SP = SP - 50
    picDMG.Print ; SP; ": spell points remaining."
    If EHP <= 0 Then
        MsgBox "You Killed the Dragon, you get 15000 treasure."
        Treasure = Treasure + 15000
        MsgBox "You Gained a level."
        Level = Level + 1
        Form15.Hide
        Form16.Show
    ElseIf EHP > 0 Then
        picDMG.Print "The Dragon does 999 damage to you."
        HP = HP - 999
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
    End If
ElseIf SP < 50 Then
    MsgBox "You do not have enough Magic points to cast this spell."
End If

End Sub

Private Sub cmdPrep_Click()
'Prepares max damage, I put a counter for the dragon to notice'
'what you do if he survives the first attack'
Dim Pos As Integer, Pass As Integer
TDMG As Integer, Preptime As Integer
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
If Preptime <= 1 Then
    picDMG.Print "The Dragon Laughs at you"
    Preptime = Preptime + 1
ElseIf Preptime >= 2 Then
    picDMG.Print "The Dragon does 90 damage to you."
    HP = HP - 90
    picDMG.Print ; HP; ": Hitpoints left"
    If HP <= 0 Then
        MsgBox "You Die. Game over. Please Try again."
        End
    End If
End If
End Sub

Private Sub cmdSneak_Click()
'Used just to be sarcastic'
MsgBox ("Dude, you can't sneak past this guy, he's a freaking dragon.")
End Sub

Private Sub cmdSurrender_Click()
'essentially an end button, that mocks you as well'
MsgBox "The Dragon laughs and eats you. You were so close too. Game Over."
End
End Sub

Private Sub Form_Load()
'Loads up everything needed for the fight.'
CTR = 0
Open Path & "fighterdmg.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, People(CTR), Bday(CTR)
    picResults.Print People(CTR), Bday(CTR)
Loop
Close
HP = 100
SP = 50
EHP = 1000
CTR = 0
picMonster.Picture = LoadPicture("Dragon.jpg")
picDMG.Print "A Dragon sees you and laughs outright."
End Sub

