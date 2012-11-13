VERSION 5.00
Begin VB.Form frmCharizardVsBlastoise 
   BackColor       =   &H80000007&
   Caption         =   "Charizard Vs Blastoise"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "frmCharizardVsBlastoise.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHpBlastoise 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox picHpCharizard 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox PicMenu 
      Height          =   1215
      Left            =   0
      Picture         =   "frmCharizardVsBlastoise.frx":4F28
      ScaleHeight     =   1155
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   3240
      Width           =   6735
      Begin VB.CommandButton cmdFireBlast 
         Caption         =   "Fire Blast"
         Height          =   615
         Left            =   2400
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdWingAttack 
         Caption         =   "Wing Attack"
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdDragonClaw 
         Caption         =   "Dragonclaw"
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdFlamethrower 
         Caption         =   "Flamethrower"
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   495
         Left            =   4560
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   735
         Left            =   5640
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdFight 
         Caption         =   "Fight"
         Height          =   735
         Left            =   4560
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picCharBack 
      Height          =   1215
      Left            =   720
      Picture         =   "frmCharizardVsBlastoise.frx":649E
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox picBlastoiseEncounter 
      Height          =   1215
      Left            =   4680
      Picture         =   "frmCharizardVsBlastoise.frx":9B92
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCharizardVsBlastoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Project
'frmCharizardVsBlastoise
'Eugene Zheng
'10/10/2009
'This is the screen the user fights Blastoise using Charizard
'One uses to 4 attack command buttons on to lower the other pokemon's HP
'We can use a random number generator to enable the computer pokemon to act on its own
'Essentially, everything is governed by If- Then Statements
'Actions are taken or not taken by the If- Then Statements

Option Explicit
Dim RndNumber As Integer


Private Sub cmdDragonClaw_Click()

'Message box to tell the user what is going on
'91 is the damage given
MsgBox "Charizard used Dragon Claw!", , "Charizard"
BlasHp = BlasHp - 91

'Print the remaining HP for each pokemon
picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp

'If the opponents HP is zero these Case statement will convey the message that the opponent is defeated and bring the user back one screen
Select Case BlasHp
    Case Is <= 0
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:0"
        MsgBox "Blastoise Fainted", , "Blastoise"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
        frmCharizardVsBlastoise.Hide
        frmCharizardChosen.Show
End Select

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If BlasHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Blastoise used Hydro Pump!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        CharHp = CharHp - CharHp
        MsgBox "It's super effective!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Blastoise used Surf!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        CharHp = CharHp - 219
        MsgBox "It's super effective!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Blastoise used Ice Beam!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        CharHp = CharHp - 22
        MsgBox "It's not very effective...", , "Charizard"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
           cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Blastoise used Skull Bash!", , "Blastoise"
    CharHp = CharHp - 69
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    
    Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        Case Else
            picHpCharizard.Cls
            picHpCharizard.Print "HP:"; CharHp
        End Select
End If
End If
End Sub

Private Sub cmdFight_Click()
CharHp = 270
BlasHp = 292


cmdFlamethrower.Visible = True
cmdDragonClaw.Visible = True
cmdWingAttack.Visible = True
cmdFireBlast.Visible = True
cmdFight.Visible = False
picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp

picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp
End Sub

Private Sub cmdFireBlast_Click()
MsgBox "Charizard used Fire Blast!", , "Charizard"
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)

If RndNumber <= 3 Then
    BlasHp = BlasHp - 37
    MsgBox "It's not very effective...", , "Charizard"
ElseIf RndNumber = 4 Then
    MsgBox "Charizard Missed", , "Charizard"
Else
    BlasHp = BlasHp - 37
    MsgBox "It's not very effective...", , "Charizard"
End If


picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp


Select Case BlasHp
    Case Is <= 0
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:0"
        MsgBox "Blastoise Fainted", , "Blastoise"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
        frmCharizardVsBlastoise.Hide
        frmCharizardChosen.Show
End Select

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If BlasHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Blastoise used Hydro Pump!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        CharHp = CharHp - CharHp
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's super effective!", , "Blastoise"
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Blastoise used Surf!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        CharHp = CharHp - 219
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's super effective!", , "Blastoise"
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Blastoise used Ice Beam!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        CharHp = CharHp - 22
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's not very effective...", , "Charizard"
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Blastoise used Skull Bash!", , "Blastoise"
    CharHp = CharHp - 69
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
        
    Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        Case Else
            picHpCharizard.Cls
            picHpCharizard.Print "HP:"; CharHp
        End Select
End If
End If
End Sub

Private Sub cmdFlamethrower_Click()

MsgBox "Charizard used Flamethrower!", , "Charizard"
MsgBox "It's not very effective...", , "Charizard"
BlasHp = BlasHp - 17

picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp


Select Case BlasHp
    Case Is <= 0
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:0"
        MsgBox "Blastoise Fainted", , "Blastoise"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
        frmCharizardVsBlastoise.Hide
        frmCharizardChosen.Show
End Select

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If BlasHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Blastoise used Hydro Pump!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        CharHp = CharHp - CharHp
        MsgBox "It's super effective!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Blastoise used Surf!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        CharHp = CharHp - 219
        MsgBox "It's super effective!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Blastoise used Ice Beam!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        CharHp = CharHp - 22
        MsgBox "It's not very effective...", , "Charizard"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Blastoise used Skull Bash!", , "Blastoise"
    CharHp = CharHp - 69
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
        
    Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
           cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        Case Else
            picHpCharizard.Cls
            picHpCharizard.Print "HP:"; CharHp
        End Select
End If
End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRun_Click()
cmdFight.Visible = True
cmdFlamethrower.Visible = False
cmdDragonClaw.Visible = False
cmdWingAttack.Visible = False
cmdFireBlast.Visible = False
frmCharizardVsBlastoise.Hide
frmCharizardChosen.Show
End Sub



Private Sub cmdWingAttack_Click()

MsgBox "Charizard used Wing Attack!", , "Charizard"
BlasHp = BlasHp - 74

picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp


Select Case BlasHp
    Case Is <= 0
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:0"
        MsgBox "Blastoise Fainted", , "Blastoise"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
        frmCharizardVsBlastoise.Hide
        frmCharizardChosen.Show
End Select

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If BlasHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Blastoise used Hydro Pump!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        CharHp = CharHp - CharHp
        MsgBox "It's super effective!", , "Blastoise"
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
           cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Blastoise used Surf!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        CharHp = CharHp - 219
        MsgBox "It's super effective!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Blastoise used Ice Beam!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        CharHp = CharHp - 22
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's not very effective...", , "Charizard"
        Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Blastoise used Skull Bash!", , "Blastoise"
    CharHp = CharHp - 69
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
        
    Select Case CharHp
        Case Is <= 0
            picHpCharizard.Cls
            picHpCharizard.Print "HP:0"
            MsgBox "Charizard Fainted", , "Charizard"
            MsgBox "Sorry, you have lost", , "Charizard"
            cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
            frmCharizardVsBlastoise.Hide
            frmCharizardChosen.Show
        Case Else
            picHpCharizard.Cls
            picHpCharizard.Print "HP:"; CharHp
        End Select
End If
End If
End Sub

