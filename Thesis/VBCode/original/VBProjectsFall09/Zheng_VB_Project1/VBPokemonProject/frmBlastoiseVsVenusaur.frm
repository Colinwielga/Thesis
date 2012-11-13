VERSION 5.00
Begin VB.Form frmBlastoiseVsVenusaur 
   Caption         =   "Blastoise Vs Venusaur"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   Picture         =   "frmBlastoiseVsVenusaur.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHpBlastoise 
      BackColor       =   &H80000004&
      Height          =   255
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   1095
      Left            =   720
      Picture         =   "frmBlastoiseVsVenusaur.frx":1F66
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.PictureBox picMenu 
      Height          =   1215
      Left            =   0
      Picture         =   "frmBlastoiseVsVenusaur.frx":2A22
      ScaleHeight     =   1155
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   3240
      Width           =   6735
      Begin VB.CommandButton cmdHydroPump 
         Caption         =   "Hydro Pump"
         Height          =   615
         Left            =   2400
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdSkullBash 
         Caption         =   "Skull Bash"
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdIceBeam 
         Caption         =   "Ice Beam"
         Height          =   615
         Left            =   2400
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdSurf 
         Caption         =   "Surf"
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
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   735
         Left            =   5640
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.PictureBox picHpVenusaur 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox picVenuEncounter 
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4680
      Picture         =   "frmBlastoiseVsVenusaur.frx":3F98
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmBlastoiseVsVenusaur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Project
'frmBlastoiseVsVenusaur
'Eugene Zheng
'10/17/ 2009
'This is the form where user uses Blastoise to fight Venusaur
'One uses to 4 attack command buttons on to lower the other pokemon's HP
'We can use a random number generator to enable the computer pokemon to act on its own
'Essentially, everything is governed by If- Then Statements
'Actions are taken or not taken by the If- Then Statements

Option Explicit
Dim RndNumber As Integer
Dim SolarBeam As Boolean


Private Sub cmdIceBeam_Click()
'If Solarbeam was chosen by the computer, we don't want the attack to be interrupted
'Thus if solarbeam is true then the computer doesn't generate another attack
'This is accomplished by using If- Then statements

If SolarBeam = True Then
    SolarBeam = False
    MsgBox "Venusaur used Solar Beam!", , "Venusaur"
    BlasHp = BlasHp - BlasHp
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:"; BlasHp
    MsgBox "It's super effective!", , "Venusaur"
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
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
        
        BlasHp = BlasHp - 270
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's super effective!", , "Venusaur"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's super effective!", , "Venusaur"
        End If
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Venusaur used Earthquake!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        BlasHp = BlasHp - 21
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's not very effective...", , "Venusaur"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's not very effective...", , "Venusaur"
        End If
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Venusaur used Take Down!", , "Venusaur"
    MsgBox "Venusaur was hit with recoil!", , "Venusaur"
    BlasHp = BlasHp - 25
    VenuHp = VenuHp - 7
    If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End If
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
    
    Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        Case Else
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
    End Select
End If
End If
End If
    
If BlasHp > 0 Then
    MsgBox "Blastoise used Ice Beam!", , "Blastoise"
VenuHp = VenuHp - 127
MsgBox "It's super effective!", , "Blastoise"

picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp
picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)

Select Case VenuHp
    Case Is <= 0
        SolarBeam = False
        picHpVenusaur.Cls
        picHpVenusaur.Print "HP:0"
        MsgBox "Venusaur Fainted", , "Venusaur"
        MsgBox "Congratulations, You Won!", , "Winner"
        cmdFight.Visible = True
        cmdSurf.Visible = False
        cmdIceBeam.Visible = False
        cmdSkullBash.Visible = False
        cmdHydroPump.Visible = False
        frmBlastoiseVsVenusaur.Hide
        frmBlastoiseChosen.Show
End Select
End If
    

End Sub

Private Sub cmdFight_Click()
BlasHp = 291
VenuHp = 309


cmdSurf.Visible = True
cmdHydroPump.Visible = True
cmdSkullBash.Visible = True
cmdIceBeam.Visible = True
cmdFight.Visible = False
picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp

picHpVenusaur.Cls
picHpVenusaur.Print "HP:"; VenuHp
End Sub

Private Sub cmdHydroPump_Click()




RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If SolarBeam = True Then
    SolarBeam = False
    MsgBox "Venusaur used Solar Beam!", , "Venusaur"
    BlasHp = BlasHp - BlasHp
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:"; BlasHp
    MsgBox "It's super effective!", , "Venusaur"
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
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
        
        BlasHp = BlasHp - 270
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's super effective!", , "Venusaur"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's super effective!", , "Venusaur"
        End If
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Venusaur used Earthquake!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        BlasHp = BlasHp - 21
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's not very effective...", , "Venusaur"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's not very effective", , "Venusaur"
        End If
        
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Venusaur used Take Down!", , "Venusaur"
    MsgBox "Venusaur was hit with recoil!", , "Venusaur"
    BlasHp = BlasHp - 25
    VenuHp = VenuHp - 7
   
   If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
           
        End If
        
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
    
    Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        Case Else
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End Select
        
End If
End If
End If
        
    If BlasHp > 0 Then
    
        MsgBox "Blastoise used Hydro Pump!", , "Blastoise"
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)

If RndNumber <= 3 Then
    VenuHp = VenuHp - 44
    MsgBox "It's not very effective!", , "Blastoise"
ElseIf RndNumber = 4 Then
    MsgBox "Blastoise Missed", , "Blastoise"
Else
    VenuHp = VenuHp - 44
    MsgBox "It's not very effective", , "Blastoise"
End If


picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp
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
        cmdSurf.Visible = False
        cmdIceBeam.Visible = False
        cmdSkullBash.Visible = False
        cmdHydroPump.Visible = False
        frmBlastoiseVsVenusaur.Hide
        frmBlastoiseChosen.Show
End Select
End If

End Sub

Private Sub cmdSurf_Click()


RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If SolarBeam = True Then
    SolarBeam = False
    MsgBox "Venusaur used Solar Beam!", , "Venusaur"
    BlasHp = BlasHp - BlasHp
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:"; BlasHp
    MsgBox "It's super effective!", , "Venusaur"
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
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
        
        BlasHp = BlasHp - 270
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's super effective!", , "Venusaur"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's super effective!", , "Venusaur"
        End If
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Venusaur used Earthquake!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        BlasHp = BlasHp - 21
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's not very effective...", , "Venusaur"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's not very effective...", , "Venusaur"
        End If
        
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Venusaur used Take Down!", , "Venusaur"
    MsgBox "Venusaur was hit with recoil!", , "Venusaur"
    BlasHp = BlasHp - 25
    VenuHp = VenuHp - 7
    If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End If
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
    
End If
End If
End If
    
    Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        Case Else
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End Select
        
    If BlasHp > 0 Then
        MsgBox "Blastoise used Surf!", , "Blastoise"
        MsgBox "It's not very effective...", , "Blastoise"
        VenuHp = VenuHp - 31

picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp
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
        cmdSurf.Visible = False
        cmdIceBeam.Visible = False
        cmdSkullBash.Visible = False
        cmdHydroPump.Visible = False
        frmBlastoiseVsVenusaur.Hide
        frmBlastoiseChosen.Show
End Select
End If

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRun_Click()
cmdFight.Visible = True
cmdSurf.Visible = False
cmdIceBeam.Visible = False
cmdSkullBash.Visible = False
cmdHydroPump.Visible = False
frmBlastoiseVsVenusaur.Hide
frmBlastoiseChosen.Show
End Sub



Private Sub cmdSkullBash_Click()



RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If SolarBeam = True Then
    SolarBeam = False
    MsgBox "Venusaur used Solar Beam!", , "Venusaur"
    
    BlasHp = BlasHp - BlasHp
    picHpBlastoise.Cls
    picHpBlastoise.Print "HP:"; BlasHp
    MsgBox "It's super effective!", , "Venusaur"
    Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
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
        BlasHp = BlasHp - 270
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's super effective!", , "Venusaur"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's super effective!", , "Venusaur"
        End If
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Venusaur used Earthquake!", , "Venusaur"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 8 Then
        BlasHp = BlasHp - 21
       If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's not very effective...", , "Venusaur"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's not very effective...", , "Venusaur"
        End If
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsVenusaur.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Venusaur Missed!", , "Venusaur"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Venusaur used Take Down!", , "Venusaur"
    MsgBox "Venusaur was hit with recoil!", , "Venusaur"
    BlasHp = BlasHp - 25
    VenuHp = VenuHp - 7

End If
End If
End If

 If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End If
    picHpVenusaur.Cls
    picHpVenusaur.Print "HP:"; VenuHp
    
   
If BlasHp > 0 Then
        MsgBox "Blastoise used Skull Bash!", , "Blastoise"
VenuHp = VenuHp - 56

picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp
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
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
        frmBlastoiseVsVenusaur.Hide
        frmBlastoiseChosen.Show
End Select
End If
        
End Sub

