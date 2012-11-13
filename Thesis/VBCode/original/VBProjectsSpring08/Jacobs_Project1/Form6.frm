VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H000000C0&
   Caption         =   "Form6"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H000000C0&
   HasDC           =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDmg 
      BackColor       =   &H000000C0&
      Height          =   2775
      Left            =   6840
      ScaleHeight     =   2715
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
   Begin VB.PictureBox picMonster 
      BackColor       =   &H000000C0&
      Height          =   2655
      Left            =   4560
      ScaleHeight     =   2595
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdPrep 
      Caption         =   "Prepare"
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdBolt 
      Caption         =   "Lightning Bolt"
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdDefend 
      Caption         =   "Defend"
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "Fight"
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SP As Integer, HP As Integer, EHP As Integer
Dim DMG(1 To 20) As Integer, CTR As Integer
Private Sub cmdBolt_Click()
'First spell damage, used similarly to how melee damage works.'
If SP >= 5 Then
    picDMG.Print "The Skeleton takes 50 dmg"
    EHP = EHP - 50
    SP = SP - 5
    picDMG.Print ; SP; ": spell points remaining."
    If EHP <= 0 Then
        MsgBox "You Killed the Skeleton, you get 100 treasure."
        Treasure = Treasure + 100
        MsgBox "You gained a level"
        Level = Level + 1
        Form6.Hide
        Form7.Show
    ElseIf EHP > 0 Then
        picDMG.Print "The Skeleton does 5 damage to you."
        HP = HP - 5
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
picDMG.Print "The Skeleton swings and misses."
End Sub

Private Sub cmdFight_Click()
'Melee Attacks'
CTR = CTR + 1
EHP = EHP - DMG(CTR)
If DMG(CTR) >= 200 Then
    picDMG.Print "CRITICAL HIT!"
ElseIf DMG(CTR) < 1 Then
    picDMG.Print "You Miss."
ElseIf DMG(CTR) = 1 - 199 Then
    picDMG.Print "You Hit the Skeleton for"; DMG(CTR); " Damage"
End If
If EHP <= 0 Then
        MsgBox "You Killed the Skeleton, you get 100 treasure."
        Treasure = Treasure + 100
        MsgBox "You gained a level."
        Level = Level + 1
        Form6.Hide
        Form7.Show
ElseIf EHP > 0 Then
        picDMG.Print "The skeleton does 4 damage to you."
        HP = HP - 4
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
End If
If CTR = 20 Then CTR = 0
End Sub

Private Sub cmdPrep_Click()
Dim Pos As Integer, Pass As Integer
Dim TDMG As Integer
'Sorts out the damage'
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
picDMG.Print "The Skeleton does 3 damage to you."

HP = HP - 3
picDMG.Print ; HP; ": Hitpoints left"
If HP <= 0 Then
    MsgBox "You Die. Game over. Please Try again."
    End
End If
End Sub


Private Sub Form_Load()
'Shows another fight and gets it ready'
CTR = 0
Open Path & "magedmg.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, DMG(CTR)
Loop
Close
HP = 75
SP = 10
EHP = 75
CTR = 5
picmonster.Picture = LoadPicture("Skeleton.jpg")
picDMG.Print "A skeleton stands up and begins to attack you."
End Sub
