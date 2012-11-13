VERSION 5.00
Begin VB.Form frmVenusaurVsBlastoise 
   Caption         =   "Venusaur Vs Blastoise"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   Picture         =   "frmVenusaurVsBlastoise.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHpVenusaur 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picMenu 
      Height          =   1215
      Left            =   0
      Picture         =   "frmVenusaurVsBlastoise.frx":4F28
      ScaleHeight     =   1155
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   3240
      Width           =   6735
      Begin VB.CommandButton cmdSolarbeam 
         Caption         =   "Solarbeam"
         Height          =   615
         Left            =   2400
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdTakeDown 
         Caption         =   "Take Down"
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdEarthquake 
         Caption         =   "Earthquake"
         Height          =   615
         Left            =   2400
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdRazorLeaf 
         Caption         =   "Razor Leaf"
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdFight 
         Caption         =   "Fight"
         Height          =   735
         Left            =   4560
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   735
         Left            =   5640
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   495
         Left            =   4560
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.PictureBox picVenusaurBack 
      BackColor       =   &H80000009&
      Height          =   1095
      Left            =   720
      Picture         =   "frmVenusaurVsBlastoise.frx":649E
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox picHpBlastoise 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox picBlastoiseEncounter 
      Height          =   1215
      Left            =   4680
      Picture         =   "frmVenusaurVsBlastoise.frx":70EC
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmVenusaurVsBlastoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Project
'frmVenusaurVsBlastoise
'Eugene Zheng
'10/17/2009
'This is the screen where the user uses Venusaur to fight Blastoise
'One uses to 4 attack command buttons on to lower the other pokemon's HP
'We can use a random number generator to enable the computer pokemon to act on its own
'Essentially, everything is governed by If- Then Statements
'Actions are taken or not taken by the If- Then Statements

Option Explicit
Dim RndNumber As Integer
Dim SolarBeam As Boolean
Dim CTR As Integer



Private Sub cmdRazorLeaf_Click()
'One of Venusaur's attack
'the value of 210 is the amoung of damage
MsgBox "Venusaur used Razor Leaf!", , "Venusaur"
MsgBox "It's super effective!", , "Venusaur"
BlasHp = BlasHp - 210

'Pring the remaing HP
If BlasHp > 0 Then
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:"; BlasHp
ElseIf BlasHp <= 0 Then
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:0"
End If

'Code to determine if the user has won
Select Case BlasHp
    Case Is <= 0
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:0"
        MsgBox "Blastoise Fainted", , "Blastoise"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
        frmVenusaurVsBlastoise.Hide
        frmVenusaurChosen.Show
End Select

'Random number generator to determine which move the computer will use
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
        VenuHp = VenuHp - 76
        MsgBox "It's not very effective...", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Blastoise used Surf!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        VenuHp = VenuHp - 43
        MsgBox "It's not very effective...", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Blastoise used Ice Beam!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        VenuHp = VenuHp - 110
        MsgBox "It's super effective!", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Blastoise used Skull Bash!", , "Blastoise"
    VenuHp = VenuHp - 73
    If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
End If
End If
End Sub

Private Sub cmdFight_Click()

'This shows the buttons necessary to fight
'This button also enables the user to fight the pokemon multiple times
VenuHp = 309
BlasHp = 292
CTR = 0

cmdRazorLeaf.Visible = True
cmdTakeDown.Visible = True
cmdEarthquake.Visible = True
cmdSolarbeam.Visible = True
cmdFight.Visible = False
picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp

picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp
End Sub

Private Sub cmdSolarbeam_Click()

If SolarBeam = False Then
    CTR = 0
    SolarBeam = True
    MsgBox "Venusaur is gathering sunlight!", , "Venusaur"
    cmdRazorLeaf.Visible = False
    cmdTakeDown.Visible = False
    cmdEarthquake.Visible = False
    CTR = CTR + 1
ElseIf SolarBeam = True And CTR = 1 Then
    MsgBox "Venusaur is still gathering sunlight!", , "Venusaur"
    CTR = CTR + 1
ElseIf SolarBeam = True And CTR = 2 Then
    SolarBeam = False
    MsgBox "Venusaur used Solarbeam!", , "Venusaur"
    BlasHp = BlasHp - BlasHp
    cmdRazorLeaf.Visible = True
    cmdTakeDown.Visible = True
    cmdEarthquake.Visible = True
    CTR = 0
End If

If BlasHp > 0 Then
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:"; BlasHp
ElseIf BlasHp <= 0 Then
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:0"
End If


Select Case BlasHp
    Case Is <= 0
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:0"
        MsgBox "Blastoise Fainted", , "Blastoise"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
        frmVenusaurVsBlastoise.Hide
        frmVenusaurChosen.Show
End Select



If BlasHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If RndNumber = 4 Then

    MsgBox "Blastoise used Hydro Pump!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        VenuHp = VenuHp - 76
        MsgBox "It's not very effective...", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Blastoise used Surf!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        VenuHp = VenuHp - 43
        MsgBox "It's not very effective...", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Blastoise used Ice Beam!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        VenuHp = VenuHp - 110
        MsgBox "It's super effective!", , "Blastoise"
       If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Blastoise used Skull Bash!", , "Blastoise"
    VenuHp = VenuHp - 74
   If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
End If
End If
End Sub

Private Sub cmdEarthquake_Click()

MsgBox "Venusaur used Earthquake!", , "Venusaur"
MsgBox "It's not very effective...", , "Venusaur"
BlasHp = BlasHp - 42

If BlasHp > 0 Then
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:"; BlasHp
ElseIf BlasHp <= 0 Then
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:0"
End If


Select Case BlasHp
    Case Is <= 0
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:0"
        MsgBox "Blastoise Fainted", , "Blastoise"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
        frmVenusaurVsBlastoise.Hide
        frmVenusaurChosen.Show
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
        VenuHp = VenuHp - 76
        MsgBox "It's not very effective...", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Blastoise used Surf!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        VenuHp = VenuHp - 43
        MsgBox "It's not very effective...", , "Blastoise"
       If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Blastoise used Ice Beam!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        VenuHp = VenuHp - 110
        MsgBox "It's super effective!", , "Blastoise"
       If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Blastoise used Skull Bash!", , "Blastoise"
    VenuHp = VenuHp - 69
    If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
End If
End If
End Sub

Private Sub cmdQuit_Click()
'A simple quit button
End
End Sub

Private Sub cmdRun_Click()
'this is essentially a back button
cmdFight.Visible = True
cmdRazorLeaf.Visible = False
cmdTakeDown.Visible = False
cmdEarthquake.Visible = False
cmdSolarbeam.Visible = False
frmVenusaurVsBlastoise.Hide
frmVenusaurChosen.Show
CTR = 0
End Sub



Private Sub cmdTakeDown_Click()

MsgBox "Venusaur used Take Down!", , "Venusaur"
MsgBox "Venusaur was hit by the recoil!", , "Venusaur"
BlasHp = BlasHp - 65
VenuHp = VenuHp - 11

If VenuHp > 0 Then
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
Else
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:0"
End If

If BlasHp > 0 Then
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:"; BlasHp
ElseIf BlasHp <= 0 Then
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:0"
End If


Select Case BlasHp
    Case Is <= 0
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:0"
        MsgBox "Blastoise Fainted", , "Blastoise"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
        frmVenusaurVsBlastoise.Hide
        frmVenusaurChosen.Show
    Case Else
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
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
        VenuHp = VenuHp - 76
        MsgBox "It's not very effective...", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Blastoise used Surf!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        VenuHp = VenuHp - 43
        MsgBox "It's not very effective...", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Blastoise used Ice Beam!", , "Blastoise"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        VenuHp = VenuHp - 110
        MsgBox "It's super effective!", , "Blastoise"
        If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Blastoise Missed!", , "Blastoise"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Blastoise used Skull Bash!", , "Blastoise"
    VenuHp = VenuHp - 71
    If VenuHp > 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:"; VenuHp
        ElseIf VenuHp <= 0 Then
            picHpVenusaur.Cls
            picHpVenusaur.Print "HP:0"
            MsgBox "Venusaur Fainted", , "Venusaur"
            MsgBox "Sorry, you have lost", , "Venusaur"
            cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
            frmVenusaurVsBlastoise.Hide
            frmVenusaurChosen.Show
        End If
End If
End If
End Sub

