VERSION 5.00
Begin VB.Form frmVenusaurVsCharizard 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   Picture         =   "frmVenusaurVsCharizard.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMenu 
      Height          =   1215
      Left            =   0
      Picture         =   "frmVenusaurVsCharizard.frx":EA49
      ScaleHeight     =   1155
      ScaleWidth      =   6675
      TabIndex        =   4
      Top             =   3240
      Width           =   6735
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   495
         Left            =   4560
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   735
         Left            =   5640
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdFight 
         Caption         =   "Fight"
         Height          =   735
         Left            =   4560
         TabIndex        =   9
         Top             =   0
         Width           =   1095
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
      Begin VB.CommandButton cmdEarthquake 
         Caption         =   "Earthquake"
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdTakeDown 
         Caption         =   "Take Down"
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdSolarbeam 
         Caption         =   "Solarbeam"
         Height          =   615
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.PictureBox picHpVenusaur 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox picVenusaurBack 
      BackColor       =   &H80000009&
      Height          =   1095
      Left            =   600
      Picture         =   "frmVenusaurVsCharizard.frx":FFBF
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.PictureBox picHpCharizard 
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      FillColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   4800
      Picture         =   "frmVenusaurVsCharizard.frx":10C0D
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmVenusaurVsCharizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Project
'frmVenusaurVsBlastoise
'Eugene Zheng
'10/17/2009
'This is the screen where the user uses Venusaur to fight Charizard
'One uses to 4 attack command buttons on to lower the other pokemon's HP
'We can use a random number generator to enable the computer pokemon to act on its own
'Essentially, everything is governed by If- Then Statements
'Actions are taken or not taken by the If- Then Statements

Option Explicit
Dim RndNumber As Integer
Dim SolarBeam As Boolean
Dim CTR As Integer



Private Sub cmdRazorLeaf_Click()


RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If CharHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Charizard used Fire Blast!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        VenuHp = VenuHp - VenuHp
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Charizard used Flamethrower!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        VenuHp = VenuHp - 217
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Charizard used Wing Attack!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        VenuHp = VenuHp - 137
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Charizard used Dragon Claw!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
End If
End If

If VenuHp > 0 Then
    MsgBox "Venusaur used Razor Leaf!", , "Venusaur"
    MsgBox "It's not very effective...", , "Venusaur"
    CharHp = CharHp - 31

    If CharHp > 0 Then
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
    ElseIf CharHp <= 0 Then
        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
    End If

Select Case CharHp
    Case Is <= 0
        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
        MsgBox "Charizard Fainted", , "Charizard"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
        frmVenusaurVsCharizard.Hide
        frmVenusaurChosen.Show
End Select
End If
End Sub

Private Sub cmdFight_Click()
VenuHp = 309
CharHp = 292
CTR = 0

SolarBeam = False

cmdRazorLeaf.Visible = True
cmdTakeDown.Visible = True
cmdEarthquake.Visible = True
cmdSolarbeam.Visible = True
cmdFight.Visible = False
picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp

picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
End Sub

Private Sub cmdSolarbeam_Click()

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If CharHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Charizard used Fire Blast!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        VenuHp = VenuHp - VenuHp
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Charizard used Flamethrower!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        VenuHp = VenuHp - 217
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Charizard used Wing Attack!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        VenuHp = VenuHp - 137
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Charizard used Dragon Claw!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
End If
End If


If VenuHp > 0 Then

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
    CharHp = CharHp - 87
    MsgBox "It's not very effective...", , "Venusaur"
    cmdRazorLeaf.Visible = True
    cmdTakeDown.Visible = True
    cmdEarthquake.Visible = True
    CTR = 0
End If

If CharHp > 0 Then
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
ElseIf CharHp <= 0 Then
    picHpCharizard.Cls
    picHpCharizard.Print "HP:0"
End If


Select Case CharHp
    Case Is <= 0
        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
        MsgBox "Charizard Fainted", , "Charizard"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
        frmVenusaurVsCharizard.Hide
        frmVenusaurChosen.Show
End Select
End If
End Sub

Private Sub cmdEarthquake_Click()

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If CharHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Charizard used Fire Blast!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        VenuHp = VenuHp - VenuHp
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 3 Then

   MsgBox "Charizard used Flamethrower!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        VenuHp = VenuHp - 217
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Charizard used Wing Attack!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        VenuHp = VenuHp - 132
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 1 Then
    MsgBox "Charizard used Dragon Claw!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
End If
End If


If VenuHp > 0 Then

MsgBox "Venusaur used Earthquake!", , "Venusaur"
MsgBox "It's super effective!", , "Venusaur"
CharHp = CharHp - 132

If CharHp > 0 Then
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
ElseIf CharHp <= 0 Then
    picHpCharizard.Cls
    picHpCharizard.Print "HP:0"
End If


Select Case CharHp
    Case Is <= 0
        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
        MsgBox "Charizard Fainted", , "Charizard"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
        frmVenusaurVsCharizard.Hide
        frmVenusaurChosen.Show
End Select
End If

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRun_Click()
cmdFight.Visible = True
cmdRazorLeaf.Visible = False
cmdTakeDown.Visible = False
cmdEarthquake.Visible = False
cmdSolarbeam.Visible = False
frmVenusaurVsCharizard.Hide
frmVenusaurChosen.Show
CTR = 0
End Sub



Private Sub cmdTakeDown_Click()



RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If CharHp > 0 Then
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

    MsgBox "Charizard used Fire Blast!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((4 - 1 + 1) * Rnd + 1)
    If RndNumber <= 3 Then
        VenuHp = VenuHp - VenuHp
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 1 Then

   MsgBox "Charizard used Flamethrower!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        VenuHp = VenuHp - 219
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Charizard used Wing Attack!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        VenuHp = VenuHp - 132
        MsgBox "It's super effective!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:"; VenuHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Charizard used Dragon Claw!", , "Charizard"
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
            frmVenusaurVsCharizard.Hide
            frmVenusaurChosen.Show
        End If
End If
End If

If VenuHp > 0 Then

MsgBox "Venusaur used Take Down!", , "Venusaur"
MsgBox "Venusaur was hit by the recoil!", , "Venusaur"
CharHp = CharHp - 92
VenuHp = VenuHp - 11

If VenuHp > 0 Then
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
Else
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:0"
End If

If CharHp > 0 Then
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
ElseIf CharHp <= 0 Then
    picHpCharizard.Cls
    picHpCharizard.Print "HP:0"
End If


Select Case CharHp
    Case Is <= 0
        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
        MsgBox "Charizard Fainted", , "Charizard"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
            cmdRazorLeaf.Visible = False
            cmdEarthquake.Visible = False
            cmdTakeDown.Visible = False
            cmdSolarbeam.Visible = False
        frmVenusaurVsCharizard.Hide
        frmVenusaurChosen.Show
    Case Else
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
End Select
End If
End Sub
