VERSION 5.00
Begin VB.Form FrmBattle 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   705
   ClientTop       =   2055
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   8295
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   480
      Picture         =   "FrmBattle.frx":0000
      ScaleHeight     =   8295
      ScaleWidth      =   7695
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H8000000C&
      Caption         =   "FIGHT"
      Height          =   735
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   8400
      Width           =   1455
   End
End
Attribute VB_Name = "FrmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
Dim userinput As Integer, playermove As String, monsteraction As Integer
Dim PlayerHealth As Integer, Monsterhealth As Integer

Dim Battle(1 To 30) As Integer, ctr As Integer

Open App.Path & "\MonsterControl.txt" For Input As #1
ctr = 0
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, Battle(ctr)
Loop
Close #1
Battlestart:
    pos = 0
    
    'do while monster health
    'do while player health
    Monsterhealth = 9
    PlayerHealth = 5
    
    Do While Monsterhealth > 0 And PlayerHealth > 0
        
        pos = pos + 1
        If pos = ctr Then
            pos = 1
        End If
        
        
        'Monster move
        If Battle(pos) = 1 Then
            MsgBox ("The monster growl's fiercely and lowers its head. It's eyes glow blood red....")
        End If
        
        If Battle(pos) = 2 Then
            MsgBox ("The monster cocks its head back and lets out a deafening roar.")
        End If
        
        If Battle(pos) = 3 Then
            MsgBox ("The monster shakes its head wildly as if in confusion.")
        End If
        monsteraction = Battle(pos)
        
        playermove = vbNullString
        'This is Prof. Rahal's code for dealing with the cancel button on the input box
        
        Do While playermove = vbNullString
            playermove = InputBox("What will you do? Please enter the number: 1-Step to the side; 2-Step Back; 3-Shoot your Gun")
                    If playermove = vbNullString Then
                        MsgBox ("Your choice must be a number!")
                    Else
                        If Val(playermove) = 1 Or Val(playermove) = 2 Or Val(playermove) = 3 Then
                           userinput = Val(playermove)
                        Else
                             MsgBox ("Your choice must be 1, 2, or 3!")
                            playermove = vbNullString
                        End If
                    End If
        Loop
        
        Select Case monsteraction
            Case Is = 1                                                         'If Monster Charges
                If userinput = 1 Then
                    MsgBox ("You stepped to the side and narrowly avoided the monster's charge!")
                End If
                If userinput = 2 Then
                    MsgBox ("The monster charges forward and rams you with its head!")
                    PlayerHealth = PlayerHealth - 1
                End If
                If userinput = 3 Then
                    MsgBox ("The monster charges forward and rams you causing your shot to miss!")
                    PlayerHealth = PlayerHealth - 1
                End If
            Case Is = 2                                                             'if monster swings its tail
                If userinput = 1 Then
                    MsgBox ("The monster whips its tail around and you stepped right into it! You are striked!")
                    PlayerHealth = PlayerHealth - 1
                End If
                If userinput = 2 Then
                    MsgBox ("You dodge the monster's tail swing by stepping backward!")
                End If
                If userinput = 3 Then
                    MsgBox ("The monster strikes you quickly with its tail causing you to miss your shot!")
                    PlayerHealth = PlayerHealth - 1
                End If
            Case Is = 3                                                             'if monster hesitates
                If userinput = 1 Then
                    MsgBox ("The monster glares at you as you step to the side.")
                End If
                If userinput = 2 Then
                    MsgBox ("The monster glares at you as you step backwards.")
                End If
                If userinput = 3 Then
                    MsgBox ("You take aim and shoot your gun. Dead on! The monster stumbles back a step.")
                    Monsterhealth = Monsterhealth - 3
                End If
        End Select
    If Monsterhealth <> 0 Then
        Select Case PlayerHealth
            Case Is = 5
                MsgBox ("Adrenline is pumping through your veins. Keep moving, don't get hit!")
            Case Is = 4
                MsgBox ("You try to keep your barrings as the last hit still ripples through your body.")
            Case Is = 3
                MsgBox ("Blood drips from your a cut on your forehead.")
            Case Is = 2
                MsgBox ("Your body burns with pain from the several strikes.")
            Case Is = 1
                MsgBox ("You can barely hold yourself up... Another hit and you're a goner...")
        End Select
    End If
        
    Loop
            
Dim Retry As Integer
If PlayerHealth = 0 Then
    MsgBox ("You have fallen.")  ' Ask if player wants to start again
    Retry = MsgBox("Will you summon the will to fight again? (If not, you will go back to the title screen)", vbYesNo)
    
    If Retry = vbYes Then
        GoTo Battlestart
    Else
        frmTitleScreen.Show
        FrmBattle.Hide
    End If
End If
   

If Monsterhealth = 0 Then
    MsgBox ("The monster reels back, howling in pain. The sheer force of its scream paralyzes you.")
    MsgBox ("The monster stumbles and falls to the floor...")
    MsgBox ("All of the lights go out and an emergency alarm activates. Secondary lights switch on and illuminates the area.")
    MsgBox ("The monster had destroyed much of the machinery in the room during the fight. It must have destroyed something important")
    MsgBox ("You must evacuate quickly!")
    
    
    frmBattleConclusion.Show
    FrmBattle.Hide
End If



End Sub

Private Sub Form_activate()
    MsgBox ("The TESTING AREA door slides open and from the darkness, a monstrous creature approaches.")
End Sub
