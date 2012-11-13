VERSION 5.00
Begin VB.Form frmBlastoiseVsCharizard 
   Caption         =   "Blastoise Vs Charizard"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   Picture         =   "frmBlastoiseVsCharizard.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHpCharizard 
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      FillColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   4800
      Picture         =   "frmBlastoiseVsCharizard.frx":EA49
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox picHpBlastoise 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox picMenu 
      Height          =   1215
      Left            =   0
      Picture         =   "frmBlastoiseVsCharizard.frx":F834
      ScaleHeight     =   1155
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   3240
      Width           =   6735
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   735
         Left            =   5640
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdFight 
         Caption         =   "Fight"
         Height          =   735
         Left            =   4560
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdSurf 
         Caption         =   "Surf"
         Height          =   615
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdIceBeam 
         Caption         =   "Ice Beam"
         Height          =   615
         Left            =   2400
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdSkullBash 
         Caption         =   "Skull Bash"
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdHydroPump 
         Caption         =   "Hydro Pump"
         Height          =   615
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   1095
      Left            =   720
      Picture         =   "frmBlastoiseVsCharizard.frx":10DAA
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "frmBlastoiseVsCharizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Battle
'frmBlastoiseVsCharizard
'Eugene Zheng
'10/18/2009
'This is the battle screen between Blastoise and Charizard
'Using the 4 command moves, the user battles the Charizard
'We can use a random number generator to enable the computer pokemon to act on its own
'Essentially, everything is governed by If- Then Statements
'Actions are taken or not taken by the If- Then Statements

Option Explicit
Dim RndNumber As Integer



Private Sub cmdIceBeam_Click()
'Since the Charizard is "faster", it attacks first while Blastoise attacks second


If CharHp > 0 Then
'Random Number generator to determine which move charizard will use
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)

'We assign each "move" with a number 1-4. Using this, we have developed a system where Charizard's attacks are random
If RndNumber = 4 Then

    MsgBox "Charizard used Fire Blast!", , "Charizard"
'Another random number generator to determine if the attack misses or connects
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

'Print the remaining HP
'We first clear the picture box to make room for the new value
picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp
    

'A different attack using the random generator
ElseIf RndNumber = 1 Then

   MsgBox "Charizard used Wing Attack!", , "Charizard"
    
    'Random generator to determine if the attack connects
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        
        BlasHp = BlasHp - 66
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End If
        
        'If Blastoise loses we need code to bring the user back to the other screen
        Select Case BlasHp
        Case Is <= 0
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "Blastoise Fainted", , "Blastoise"
            MsgBox "Sorry, you have lost", , "Blastoise"
            'These buttons visibility is changed because cmdFight resets the value of the HP for the pokemon.
            'This enables the user to "Fight" against the other pokemon another time
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        End Select
    'If the attack doesn't connect
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 2 Then

    'Another attack using the random number generator
  MsgBox "Charizard used Flamethrower!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        BlasHp = BlasHp - 21
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's not very effective...", , "Charizard"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's not very effective...", , "Charizard"
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Charizard used Dragon Claw!", , "Charizard"
    BlasHp = BlasHp - 125
    If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End If
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        Case Else
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
    End Select
End If
End If

    
If BlasHp > 0 Then
    MsgBox "Blastoise used Ice Beam!", , "Blastoise"
CharHp = CharHp - 56
MsgBox "It's not very effective...", , "Blastoise"


Select Case CharHp
    Case Is <= 0
      
        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
        MsgBox "Charizard Fainted", , "Charizard"
        MsgBox "Congratulations, You Won!", , "Winner"
                    cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
        frmBlastoiseVsCharizard.Hide
        frmBlastoiseChosen.Show
    Case Else
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
End Select
End If
    

End Sub

Private Sub cmdFight_Click()
'This button enables to the user to battle against the other pokemon
'By using the visibility feature for buttons, this button enables the user to fight against the pokemon multiple times without exiting the program
BlasHp = 291
CharHp = 270


cmdSurf.Visible = True
cmdHydroPump.Visible = True
cmdSkullBash.Visible = True
cmdIceBeam.Visible = True
cmdFight.Visible = False
picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp

picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp
End Sub

Private Sub cmdHydroPump_Click()

'Another attack


RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)





If CharHp > 0 Then
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

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
    
    
ElseIf RndNumber = 1 Then

   MsgBox "Charizard used Wing Attack!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        
        BlasHp = BlasHp - 66
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Charizard used Flamethrower!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        BlasHp = BlasHp - 21
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's not very effective...", , "Charizard"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's not very effective...", , "Charizard"
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Charizard used Dragon Claw!", , "Charizard"
    BlasHp = BlasHp - 125
   
   If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
           
        End If
        
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        Case Else
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End Select
        
End If
End If
        
    If BlasHp > 0 Then
    
        MsgBox "Blastoise used Hydro Pump!", , "Blastoise"
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)

If RndNumber <= 3 Then
    CharHp = CharHp - CharHp
    MsgBox "It's super effective!", , "Blastoise"
ElseIf RndNumber = 4 Then
    MsgBox "Blastoise Missed", , "Blastoise"
Else
    CharHp = CharHp - CharHp
    MsgBox "It's super effective!", , "Blastoise"
End If

        
        Select Case CharHp
    Case Is <= 0
 
        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
        MsgBox "Charizard Fainted", , "Charizard"
        MsgBox "Congratulations, You Won!", , "Winner"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
        frmBlastoiseVsCharizard.Hide
        frmBlastoiseChosen.Show
    Case Else
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
End Select
End If

End Sub

Private Sub cmdSurf_Click()
'Another attack


RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)


If CharHp > 0 Then
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

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
    
    
ElseIf RndNumber = 1 Then

   MsgBox "Charizard used Wing Attack!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        
        BlasHp = BlasHp - 66
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Charizard used Flamethrower!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        BlasHp = BlasHp - 21
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's not very effective...", , "Charizard"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's not very effective...", , "Charizard"
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Charizard used Dragon Claw!", , "Charizard"
    BlasHp = BlasHp - 125
    
    If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End If
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        Case Else
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End Select
        
    If BlasHp > 0 Then
        MsgBox "Blastoise used Surf!", , "Blastoise"
        CharHp = CharHp - 237
        MsgBox "It's super effective!", , "Blastoise"



Select Case CharHp
    Case Is <= 0
        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
        MsgBox "Charizard Fainted", , "Charizard"
        MsgBox "Congratulations, You Won!", , "Winner"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
        frmBlastoiseVsCharizard.Hide
        frmBlastoiseChosen.Show
    Case Else
        picHpCharizard.Cls
        picHpCharizard.Print "HP:"; CharHp
End Select
End If

End Sub

Private Sub cmdQuit_Click()
'Simple quit button
End
End Sub

Private Sub cmdRun_Click()

'This button is essentially a back button
cmdFight.Visible = True
cmdSurf.Visible = False
cmdIceBeam.Visible = False
cmdSkullBash.Visible = False
cmdHydroPump.Visible = False
frmBlastoiseVsCharizard.Hide
frmBlastoiseChosen.Show
End Sub



Private Sub cmdSkullBash_Click()

'Another attack

RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)



If CharHp > 0 Then
RndNumber = 0
Randomize
RndNumber = Int((4 - 1 + 1) * Rnd + 1)
'use random number'
'assign a move a number and use random syntax to find out which move is used
If RndNumber = 4 Then

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
    
    
ElseIf RndNumber = 1 Then

   MsgBox "Charizard used WIng Attack!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        BlasHp = BlasHp - 66
        If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 2 Then

  MsgBox "Charizard used Flamethrower!", , "Charizard"
    
    RndNumber = 0
    Randomize
    RndNumber = Int((10 - 1 + 1) * Rnd + 1)
    If RndNumber <= 9 Then
        BlasHp = BlasHp - 21
       If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
            MsgBox "It's not very effective...", , "Charizard"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
            MsgBox "It's not very effective...", , "Charizard"
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        End Select
    Else: MsgBox "Charizard Missed!", , "Charizard"
        picHpBlastoise.Cls
        picHpBlastoise.Print "HP:"; BlasHp
    End If
    
ElseIf RndNumber = 3 Then
    MsgBox "Charizard used Dragon Claw!", , "Charizard"

    BlasHp = BlasHp - 125
 

End If
End If

 If BlasHp <= 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:0"
        ElseIf BlasHp > 0 Then
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End If
    picHpCharizard.Cls
    picHpCharizard.Print "HP:"; CharHp
    
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
            frmBlastoiseVsCharizard.Hide
            frmBlastoiseChosen.Show
        Case Else
            picHpBlastoise.Cls
            picHpBlastoise.Print "HP:"; BlasHp
        End Select
        
        
If BlasHp > 0 Then
        MsgBox "Blastoise used Skull Bash!", , "Blastoise"
CharHp = CharHp - 56

picHpBlastoise.Cls
picHpBlastoise.Print "HP:"; BlasHp
picHpCharizard.Cls
picHpCharizard.Print "HP:"; CharHp



Select Case CharHp
    Case Is <= 0

        picHpCharizard.Cls
        picHpCharizard.Print "HP:0"
        MsgBox "Charizard Fainted", , "Charizard"
        MsgBox "Congratulations, You Won!", , "Winner"
            cmdFight.Visible = True
            cmdSurf.Visible = False
            cmdIceBeam.Visible = False
            cmdSkullBash.Visible = False
            cmdHydroPump.Visible = False
        frmBlastoiseVsCharizard.Hide
        frmBlastoiseChosen.Show
End Select
End If
        
End Sub

