VERSION 5.00
Begin VB.Form frmCharizardVsVenusaur 
   Caption         =   "Charizard Vs Venusaur"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   Picture         =   "frmCharizardVsVenusaur.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHpVenusaur 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox picVenuEncounter 
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4680
      Picture         =   "frmCharizardVsVenusaur.frx":1F66
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox picHpCharizard 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox picCharBack 
      Height          =   1215
      Left            =   720
      Picture         =   "frmCharizardVsVenusaur.frx":2CA9
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.PictureBox picMenu 
      Height          =   1215
      Left            =   0
      Picture         =   "frmCharizardVsVenusaur.frx":639D
      ScaleHeight     =   1155
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   3600
      Width           =   6735
      Begin VB.CommandButton cmdFireBlast 
         Caption         =   "Fire Blast"
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdWingAttack 
         Caption         =   "Wing Attack"
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdDragonClaw 
         Caption         =   "Dragonclaw"
         Height          =   615
         Left            =   2400
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdFlamethrower 
         Caption         =   "Flamethrower"
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   735
         Left            =   5640
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdFight 
         Caption         =   "Fight"
         Height          =   735
         Left            =   4560
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCharizardVsVenusaur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Project
'frmCharizardVsBlastoise
'Eugene Zheng
'10/10/2009
'This is the screen where the user uses Charizard to fight Venusur
'One uses to 4 attack command buttons on to lower the other pokemon's HP
'We can use a random number generator to enable the computer pokemon to act on its own
'Essentially, everything is governed by If- Then Statements
'Actions are taken or not taken by the If- Then Statements

Option Explicit
Dim RndNumber As Integer
Dim SolarBeam As Boolean



Private Sub cmdDragonClaw_Click()

'The 48 is the damage given
MsgBox "Charizard used Dragon Claw!", , "Charizard"
VenuHp = VenuHp - 48

'Print the remaing Hp for the pokemon
picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp

'Random Number generator for the computer to decide what move to use
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)

'Select Case statement to determine if the user has won or not
Select Case VenuHp
    Case Is <= 0
        SolarBeam = False
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:0"
        MsgBox "Venusaur Fainted", , "Venusaur"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
        frmCharizardVsVenusaur.Hide
        frmCharizardChosen.Show
End Select

'Before the computer trys to use another attack we see with Venusaur needs to use Solar beam or not
If SolarBeam = True Then
    SolarBeam = False
    MsgBox "Venusaur used Solar Beam!", , "Venusaur"
    CharHp = CharHp - 55
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    MsgBox "It's not very effective...", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
Else

'Random Number generator
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If VenuHp > 0 Then
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)
'use random number'
'assign a move a number and use random syntax to find out which move is used

'one attack from the computer
If RndNumber = 4 Then

    MsgBox "Venusaur is gathering sunlight!", , "Venusaur"
    SolarBeam = True
    
'Another attack from the computer
ElseIf RndNumber = 1 Then

   MsgBox "Venusaur used Razor Leaf!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    
    'Using the random number generator we determine if the attack from Venusaur connects or not
    If RndNumber <= 9 Then
        CharHp = CharHp - 23
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's not very effective...", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
'Another attack from the computer
ElseIf RndNumber = 2 Then

  MsgBox "Venusaur used Earthquake!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        CharHp = CharHp - 170
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's super effective!", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Venusaur used Take Down!", , "Venusaur"
    MsgBox "Venusaur was hit with recoil!", , "Venusaur"
    CharHp = CharHp - 87
    VenuHp = VenuHp - 11
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
    
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        Case Else
            picHpCharizard.Cls
            picHpCharizard.Print "HP:"; CharHp
        End Select
End If
End If
End If
End Sub

Private Sub cmdFight_Click()
CharHp = 270
VenuHp = 309


cmdFlamethrower.Visible = True
cmdDragonClaw.Visible = True
cmdWingAttack.Visible = True
cmdFireBlast.Visible = True
cmdFight.Visible = False
picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp

picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp
End Sub

Private Sub cmdFireBlast_Click()
MsgBox "Charizard used Fire Blast!", , "Charizard"
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)

If RndNumber <= 3 Then
    VenuHp = VenuHp - VenuHp
    MsgBox "It's super effective!", , "Charizard"
ElseIf RndNumber = 4 Then
    MsgBox "Charizard Missed", , "Charizard"
Else
    VenuHp = VenuHp - VenuHp
    
    MsgBox "It's super effective!", , "Charizard"
End If


picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp


Select Case VenuHp
    Case Is <= 0
        SolarBeam = False
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:0"
        MsgBox "Venusaur Fainted", , "Venusaur"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
        frmCharizardVsVenusaur.Hide
        frmCharizardChosen.Show
End Select

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If SolarBeam = True Then
    SolarBeam = False
    MsgBox "Venusaur used Solar Beam!", , "Venusaur"
    CharHp = CharHp - 55
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    MsgBox "It's not very effective...", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
Else

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If VenuHp > 0 Then
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Venusaur is gathering sunlight!", , "Venusaur"
    SolarBeam = True
    
    
ElseIf RndNumber = 1 Then

   MsgBox "Venusaur used Razor Leaf!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        
        CharHp = CharHp - 23
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's not very effective...", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Venusaur used Earthquake!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        CharHp = CharHp - 168
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's super effective!", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Venusaur used Take Down!", , "Venusaur"
    MsgBox "Venusaur was hit with recoil!", , "Venusaur"
    CharHp = CharHp - 88
    VenuHp = VenuHp - 11
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
    
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        Case Else
            picHpCharizard.Cls
            picHpCharizard.Print "HP:"; CharHp
        End Select
End If
End If
End If
End Sub

Private Sub cmdFlamethrower_Click()

MsgBox "Charizard used Flamethrower!", , "Charizard"
MsgBox "It's super effective!", , "Charizard"
VenuHp = VenuHp - 211

picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp


Select Case VenuHp
    Case Is <= 0
        SolarBeam = False
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:0"
        MsgBox "Venusaur Fainted", , "Venusaur"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
        frmCharizardVsVenusaur.Hide
        frmCharizardChosen.Show
End Select

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If SolarBeam = True Then
    SolarBeam = False
    MsgBox "Venusaur used Solar Beam!", , "Venusaur"
    CharHp = CharHp - 55
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    MsgBox "It's not very effective...", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
Else

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If VenuHp > 0 Then
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Venusaur is gathering sunlight!", , "Venusaur"
    SolarBeam = True
    
    
ElseIf RndNumber = 1 Then

   MsgBox "Venusaur used Razor Leaf!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        
        CharHp = CharHp - 23
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's not very effective...", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Venusaur used Earthquake!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        CharHp = CharHp - 176
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's super effective!", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Venusaur used Take Down!", , "Venusaur"
    MsgBox "Venusaur was hit with recoil!", , "Venusaur"
    CharHp = CharHp - 89
    VenuHp = VenuHp - 11
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
    
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        Case Else
            picHpCharizard.Cls
            picHpCharizard.Print "HP:"; CharHp
        End Select
End If
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

frmCharizardVsVenusaur.Hide
frmCharizardChosen.Show
End Sub



Private Sub cmdWingAttack_Click()

MsgBox "Charizard used Wing Attack!", , "Charizard"
MsgBox "It's super effective!", , "Charizard"
VenuHp = VenuHp - 96

picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp


Select Case VenuHp
    Case Is <= 0
        SolarBeam = False
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:0"
        MsgBox "Venusaur Fainted", , "Venusaur"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdFlamethrower.Visible = False
            cmdWingAttack.Visible = False
            cmdDragonClaw.Visible = False
            cmdFireBlast.Visible = False
        frmCharizardVsVenusaur.Hide
        frmCharizardChosen.Show
End Select

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If SolarBeam = True Then
    SolarBeam = False
    MsgBox "Venusaur used Solar Beam!", , "Venusaur"
    
    CharHp = CharHp - 55
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    MsgBox "It's not very effective...", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
Else

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If VenuHp > 0 Then
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Venusaur is gathering sunlight!", , "Venusaur"
    SolarBeam = True
    
    
ElseIf RndNumber = 1 Then

   MsgBox "Venusaur used Razor Leaf!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        CharHp = CharHp - 23
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's not very effective...", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Venusaur used Earthquake!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        CharHp = CharHp - 176
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
        MsgBox "It's super effective!", , "Venusaur"
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Venusaur used Take Down!", , "Venusaur"
    MsgBox "Venusaur was hit with recoil!", , "Venusaur"
    CharHp = CharHp - 98
    VenuHp = VenuHp - 11
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
    
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
            frmCharizardVsVenusaur.Hide
            frmCharizardChosen.Show
        Case Else
            picHpCharizard.Cls
            picHpCharizard.Print "HP:"; CharHp
        End Select
End If
End If
End If
End Sub


