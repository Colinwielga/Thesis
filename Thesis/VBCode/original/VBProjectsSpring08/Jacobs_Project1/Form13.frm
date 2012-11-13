VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H000000C0&
   Caption         =   "Form13"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form13"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrep 
      Caption         =   "MIGHTY SWING"
      Height          =   975
      Left            =   600
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "Fireball"
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdBolt 
      Caption         =   "Lightning Bolt"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdDefend 
      Caption         =   "Defend"
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "Fight"
      Height          =   1095
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.PictureBox picDmg 
      BackColor       =   &H000000C0&
      Height          =   3615
      Left            =   5520
      ScaleHeight     =   3555
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   4440
      Width           =   3615
   End
   Begin VB.PictureBox picMonster 
      BackColor       =   &H000000C0&
      Height          =   3495
      Left            =   4920
      ScaleHeight     =   3435
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SP As Integer, HP As Integer, EHP As Integer
Dim DMG(1 To 20) As Integer, CTR As Integer
Private Sub cmdBolt_Click()
'spell damage'
If SP >= 10 Then
    picDMG.Print "The Ogre takes 150 dmg"
    EHP = EHP - 50
    SP = SP - 10
    picDMG.Print ; SP; ": spell points remaining."
    If EHP <= 0 Then
        MsgBox "You Killed the Ogre, you get 450 treasure."
        Treasure = Treasure + 450
        MsgBox "You Gained a Level"
        Level = Level + 1
        Form13.Hide
        Form14.Show
    ElseIf EHP > 0 Then
        picDMG.Print "The Ogre does 20 damage to you."
        HP = HP - 20
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
    End If
ElseIf SP < 10 Then
    MsgBox "You do not have enough Magic points to cast this spell."
End If
End Sub

Private Sub cmdDefend_Click()
'Defensive action, the Ogre still hits though'
picDMG.Print "You take a defensive position."
picDMG.Print "The Ogre swings and hits hard."
HP = HP - 50
picDMG.Print ; HP; ": Hitpoints left"
If HP <= 0 Then
    MsgBox "You Die. Game over. Please Try again."
    End
End If
End Sub

Private Sub cmdFight_Click()
'Melee damage function'
CTR = CTR + 1
EHP = EHP - DMG(CTR)
If DMG(CTR) >= 200 Then
    picDMG.Print "CRITICAL HIT!"
ElseIf DMG(CTR) < 1 Then
    picDMG.Print "You Miss."
ElseIf DMG(CTR) = 1 - 199 Then
    picDMG.Print "You Hit the Ogre for"; DMG(CTR); " Damage"
End If
If EHP <= 0 Then
        MsgBox "You Killed the Ogre, you get 450 treasure."
        Treasure = Treasure + 450
        MsgBox "You Gained a Level"
        Level = Level + 1
        Form13.Hide
        Form14.Show
ElseIf EHP > 0 Then
        picDMG.Print "The Ogre does 20 damage to you."
        HP = HP - 20
        picDMG.Print ; HP; ": Hitpoints left"
        If HP <= 0 Then
            MsgBox "You Die. Game over. Please Try again."
            End
        End If
End If
If CTR = 20 Then CTR = 1
End Sub

Private Sub cmdFire_Click()
'Spell Damage'
If SP >= 20 Then
    picDMG.Print "The Ogre takes 200 dmg"
    EHP = EHP - 200
    SP = SP - 20
    picDMG.Print ; SP; ": spell points remaining."
    If EHP <= 0 Then
        MsgBox "You Killed the Ogre, you get 450 treasure."
        Treasure = Treasure + 450
        Form13.Hide
        Form14.Show
    ElseIf EHP > 0 Then
        picDMG.Print "The Ogre does 20 damage to you."
        HP = HP - 20
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


Private Sub cmdPrep_Click()
'this is used to put a total of the damage array into one big hit.'
Dim totDMG As Integer

CTR = 0

Do While CTR < 21
    CTR = CTR + 1
    totDMG = totDMG + DMG(CTR)
Loop

picDMG.Print "The Ogre does 20 damage to you."
HP = HP - 20
picDMG.Print ; HP; ": Hitpoints left"
If HP <= 0 Then
    MsgBox "You Die. Game over. Please Try again."
    End
End If
picDMG.Print "You Swing with all of your might and deal "; totDMG; " points of damage."
EHP = EHP - totDMG
If EHP <= 0 Then
    MsgBox "You Killed the Ogre, you get 450 treasure."
    Treasure = Treasure + 450
    MsgBox "You Gained a Level"
    Level = Level + 1
    Form13.Hide
    Form14.Show
End If

End Sub


Private Sub Form_Load()
'This is another loader for damage, and pictures, and variables.'
CTR = 0
Open Path & "theifdmg.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, People(CTR), Bday(CTR)
    picResults.Print People(CTR), Bday(CTR)
Loop
Close
HP = 80
SP = 35
EHP = 300
CTR = 0
picMonster.Picture = LoadPicture("Ogre.jpg")
picDMG.Print "A Ogre sees you and picks up it's club."
End Sub
