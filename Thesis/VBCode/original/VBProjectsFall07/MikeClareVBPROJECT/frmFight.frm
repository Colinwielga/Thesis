VERSION 5.00
Begin VB.Form frmFight 
   Caption         =   "ALIEN BATTLE!"
   ClientHeight    =   12525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   Picture         =   "frmFight.frx":0000
   ScaleHeight     =   12525
   ScaleWidth      =   13425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue without healing..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5520
      TabIndex        =   11
      Top             =   10800
      Width           =   1695
   End
   Begin VB.CommandButton cmdHealMe 
      Caption         =   "HEAL ME!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2280
      TabIndex        =   10
      Top             =   10800
      Width           =   1695
   End
   Begin VB.TextBox txtHeal 
      Enabled         =   0   'False
      Height          =   855
      Left            =   600
      TabIndex        =   8
      Top             =   11400
      Width           =   1695
   End
   Begin VB.CommandButton cmdHeal 
      Caption         =   "CLICK WHEN FIGHT ENDS!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   7
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "CLICK TO FIGHT!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000010&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   8115
      TabIndex        =   5
      Top             =   4080
      Width           =   8175
   End
   Begin VB.CommandButton cmdDefend 
      Caption         =   "DEFEND!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      TabIndex        =   4
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "RUN!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      TabIndex        =   3
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdAttack 
      Caption         =   "ATTACK!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   2
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblheal 
      Alignment       =   2  'Center
      Caption         =   "How much will you spend to heal?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   9
      Top             =   10800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You know how to fight now, but keep in mind this alien is certainly much stronger than the last!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1455
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALIEN BATTLE!"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmFight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AlienHP As Single
Dim AlienAttack As Single
Dim AlienReward As Single
Dim Heal As Single
Dim AtkCounter As Integer

Private Sub cmdAttack_Click()   'fight 2 with the 2nd alien!  you can attack, but he will attack after.  if he dies, you proceed, if you die, game over.
    AlienHP = AlienHP - Attack
    
    picResults.Cls
    
    If AlienHP <= 0 Then
        picResults.Cls
        picResults.Print "You attacked the alien with "; Attack; " attack points."
        picResults.Print "You've killed the alien, well done!  A true alien slayer!"
        Money = Money + AlienReward
        
        picResults.Print "You got $700 for killing this alien, well done."
        picResults.Print "You now have "; FormatCurrency(Money)
        
        cmdAttack.Enabled = False
        cmdDefend.Enabled = False
        cmdRun.Enabled = False
        cmdHeal.Enabled = True
        
        
    Else
        
        picResults.Print "Nice hit, you hurt him with "; Attack; " attack points.  He now has "; AlienHP; " H.P., but he looks pissed!"
        HP = HP - AlienAttack
        picResults.Print "He attacked you with a claw!  Ouch, that was "; AlienAttack; " atack points and it brought you to "; HP; " H.P."
        picResults.Print "Make your next move."
        
    End If
    If HP <= 0 Then
        MsgBox ("You were killed by the aliens, now the world will fall!"), , ("Game Over!")
        End
    End If
        AtkCounter = AtkCounter - 1     'attack count is at 7 now, 7 attacks till he runs
    If AtkCounter <= 0 And AlienHP > 0 Then
        MsgBox ("This alien has had enough of you!  He's running away!"), , ("He's running")
        frmAlley.Hide
        frmtunnel.Show
        MsgBox ("you see a tunnel and enter."), , ("Continuing...")
    End If
End Sub

Private Sub cmdContinue_Click() 'continue to the tunnel form
 MsgBox ("you see a tunnel and enter."), , ("Continuing...")
        frmFight.Hide
        frmtunnel.Show
End Sub

Private Sub cmdDefend_Click()   'defend the attack and it only takes 15 h.p.  if you die, game over, if he runs out of attacks, you proceed to the tunnel
    AlienAttack = 15
    HP = HP - AlienAttack
    picResults.Cls
    picResults.Print "You choose to defend from his attack!"
    picResults.Print "He attacks, but you block it and only take 15 damage!"
    If HP <= 0 Then
        MsgBox ("You were killed in battle!  Now who can save the world?!"), , ("Game Over!")
        End
    Else
        picResults.Print "Your H.P. is at "; HP; " now."
        picResults.Print "Make your next move."
    End If
        AtkCounter = AtkCounter - 1
    If AtkCounter <= 0 Then
        MsgBox ("This alien has had enough of you!  He's running away!"), , ("He's running!")
        frmFight.Hide
        frmtunnel.Show
        MsgBox ("you see a tunnel and enter."), , ("Continuing...")
    End If
End Sub

Private Sub cmdFight_Click()    'load file aliens.txt.  set up the array.  set the alien's attack, h.p. and reward for winning.
    Open App.Path & "\aliens.txt" For Input As #2

    CTR = 0
    
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, aliensHP(CTR), aliensMoney(CTR), aliensAttack(CTR)
        Loop
    Close #2
    
    For Pos = 1 To CTR

        AlienHP = aliensHP(2)
        AlienAttack = aliensAttack(2)
        AlienReward = aliensMoney(2)
        AtkCounter = 7
    Next Pos

        MsgBox ("Wait a minute...This is no ordinary alien.  He is asking you to answer 3 questions!"), , ("...what the?")  'he quizes you first, if you get the quiz right, you dont fight, otherwise you do
        MsgBox ("This alien is handing you a quiz, he says if you get all three questions right, he will let you go, otherwise, he'll fight you!"), , ("A quiz...")

        frmQuiz.Show
        frmFight.Hide


        cmdAttack.Enabled = True
        cmdDefend.Enabled = True
        cmdFight.Enabled = False
        cmdRun.Enabled = True
End Sub

Private Sub cmdHeal_Click() 'clicked after battle is won.  lets you heal yourself if needed
    MsgBox ("Good work killing the alien!  Now you have the option to heal yourself."), , ("Good Job")
    MsgBox ("Enter in the text box how much you will spend on healing yourself.  $5.00 is worth 1 H.P."), , ("How much will you spend?")
    MsgBox ("Keep in mind, you have " & FormatCurrency(Money)), , (FormatCurrency(Money))
    txtHeal.Enabled = True
    lblheal.Enabled = True
    cmdHealMe.Enabled = True
    cmdContinue.Enabled = True
End Sub

Private Sub cmdHealMe_Click()   'heals 1 h.p. for $5.  enter amount in text box.  then continue into tunnel
 Heal = txtHeal.Text
    If Heal > 0 Then
        HP = HP + (Heal * (1 / 5))
        MsgBox ("You spent " & FormatCurrency(Heal) & " on healing yourself and you have " & HP & " H.P. now."), , ("H.P. = " & HP)
        MsgBox ("you see a tunnel and enter."), , ("Continuing...")
        frmFight.Hide
        frmtunnel.Show
    Else
        MsgBox ("you see a tunnel and enter."), , ("Continuing...")
        frmFight.Hide
        frmtunnel.Show
    End If
End Sub

Private Sub cmdRun_Click()      'run away if runCounter doesn't = 0, then proceed to tunnel
    RunCounter = RunCounter - 1
    If RunCounter = 0 Then
        MsgBox ("You successfully ran away!"), , ("Run Away!")
        MsgBox ("What a pansy..."), , ("Pansy...")
        frmFight.Hide
        frmtunnel.Show
        MsgBox ("you see a tunnel and enter."), , ("Continuing...")
    Else
        picResults.Print "You can't run anymore!  You must fight!"
    End If
End Sub
