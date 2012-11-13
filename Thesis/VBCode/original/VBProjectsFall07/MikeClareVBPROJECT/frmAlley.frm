VERSION 5.00
Begin VB.Form frmAlley 
   Caption         =   "In the Alley..."
   ClientHeight    =   11340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   Picture         =   "frmAlley.frx":0000
   ScaleHeight     =   11340
   ScaleWidth      =   15315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoheal 
      Caption         =   "...Or click here if you don't need to be healed."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   13560
      TabIndex        =   14
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton cmdHealMe 
      Caption         =   "HEAL ME!"
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
      Height          =   615
      Left            =   11280
      TabIndex        =   13
      Top             =   10560
      Width           =   2055
   End
   Begin VB.TextBox txtHeal 
      Enabled         =   0   'False
      Height          =   975
      Left            =   11280
      TabIndex        =   11
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton cmdheal 
      Caption         =   "CLICK WHEN THE FIGHT ENDS!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12240
      TabIndex        =   10
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "Click when you're ready to fight!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1680
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3495
      Left            =   5400
      ScaleHeight     =   3495
      ScaleWidth      =   9735
      TabIndex        =   7
      Top             =   840
      Width           =   9735
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "RUN!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2520
      TabIndex        =   6
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton cmdDefend 
      Caption         =   "DEFEND!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1680
      TabIndex        =   5
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdAttack 
      Caption         =   "ATTACK!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblHeal 
      Alignment       =   2  'Center
      Caption         =   "How much will you spend on healing?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   12
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Battle Highlights!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   9
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "This alien is young and weak.  If able, attack him, but don't underestimate his power!"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   3
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "ALIEN STATISTICS:"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"frmAlley.frx":12C17
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALIEN BATTLE!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmAlley"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AlienHP As Single
Dim AlienAttack As Single
Dim AlienReward As Single
Dim AtkCounter As Integer
Dim Heal As Single
Dim store As String



Private Sub cmdAttack_Click()   'attack the alien with your attack points.  he then attacks you.
                                'if you die, game over, if he dies, you progress
    AlienHP = AlienHP - Attack
    
    picResults.Cls
    
    If AlienHP <= 0 Then
        picResults.Cls
        picResults.Print "You attacked the alien with "; Attack; " attack points."
        picResults.Print "You've killed the alien, well done!  A true alien slayer!"
        Money = Money + AlienReward
        
        picResults.Print "You got $350 for killing this alien, well done."
        picResults.Print "You now have "; FormatCurrency(Money)
        
        cmdAttack.Enabled = False
        cmdDefend.Enabled = False
        cmdRun.Enabled = False
        cmdHeal.Enabled = True
        
        
    Else
        
        picResults.Print "Nice hit, you hurt him with "; Attack; " attack points.  He now has "; AlienHP; " H.P., but he looks pissed!"
        HP = HP - AlienAttack
        picResults.Print "He attacked you with a claw for "; AlienAttack; " attack points!  Ouch, that hurt and brought you to "; HP; " H.P."
        picResults.Print "Make your next move."
        
    End If
    
    If HP <= 0 Then
        MsgBox ("You were killed by the aliens, now the world will fall!"), , ("Game Over!")
        End
    End If
        AtkCounter = AtkCounter - 1             'he will only attack 4 times, then run
        If AtkCounter <= 0 And AlienHP > 0 Then
            MsgBox ("This alien has had enough of you!  He's running away!"), , ("He's running")
            frmAlley.Hide
            frmStreet2.Show
            MsgBox ("Okay, back in the street...now you continue further into the rubble."), , ("Continuing...")
        End If
    End Sub
    
    Private Sub cmdDefend_Click()   'defending the alien makes his attack worth 10.  if you die, game ends.  if you live he has 4 attack until running, then proceed to street2
        AlienAttack = 10
        HP = HP - AlienAttack
        picResults.Cls
        picResults.Print "You choose to defend from his attack!"
        picResults.Print "He attacks, but you block it and only take 10 damage!"
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
            frmAlley.Hide
            frmStreet2.Show
            MsgBox ("Okay, back in the street...now you continue further into the rubble."), , ("Continuing...")
        End If
    End Sub
    
    Private Sub cmdFight_Click()            'load a file aliens.txt and make the array.  set the alien's attack points and HP
    
    Open App.Path & "\aliens.txt" For Input As #2
    
    CTR = 0
        
        Do While Not EOF(2)
            CTR = CTR + 1
            Input #2, aliensHP(CTR), aliensMoney(CTR), aliensAttack(CTR)
        Loop
        Close #2
        
    For Pos = 1 To CTR
    
        AlienHP = aliensHP(1)
        AlienAttack = aliensAttack(1)
        AlienReward = aliensMoney(1)
        RunCounter = 1
        AtkCounter = 4
    Next Pos
    
        cmdAttack.Enabled = True
        cmdDefend.Enabled = True
        cmdFight.Enabled = False
        cmdRun.Enabled = True
    End Sub
    
    Private Sub cmdHeal_Click() 'after you win you can heal.
        MsgBox ("Good work killing the alien!  Now you have the option to heal yourself."), , ("Good Job")
        MsgBox ("Enter in the text box how much you will spend on healing yourself.  $5.00 is worth 1 H.P."), , ("How much will you spend?")
        MsgBox ("Keep in mind, you have " & FormatCurrency(Money)), , (FormatCurrency(Money))
        txtHeal.Enabled = True
        lblheal.Enabled = True
        cmdHealMe.Enabled = True
        cmdNoheal.Enabled = True
        
    End Sub
    
    Private Sub cmdHealMe_Click()   '$5 is 1 h.p. you can heal by typing a dollar amount in the text box.  then proceed to the street or store
        Heal = txtHeal.Text
        If Heal > 0 Then
            HP = HP + (Heal * (1 / 5))
            MsgBox ("You spent " & FormatCurrency(Heal) & " on healing yourself and you have " & HP & " H.P. now."), , ("H.P. = " & HP)
            MsgBox ("You decide to get out of the alley."), , ("Back to the street...")
            frmAlley.Hide
            frmStreet2.Show
        Else
            MsgBox ("You decide to get out of the alley."), , ("Back to the street...")
            frmAlley.Hide
            frmStreet2.Show
        End If
        MsgBox ("Do you want to visit the old man's store again?"), , ("Store?")
          Do Until store = "yes" Or store = "no"
                store = InputBox("Type 'yes' to go to the store, or 'no' to continue...", "Yes or No?")
                If store = "yes" Then
                    frmStore2.Show
                    frmStreet2.Hide
                ElseIf store = "no" Then
                    frmStreet2.Show
                Else
                    MsgBox ("invalid entry."), , ("Error")
                End If
           Loop
    End Sub
    
    Private Sub cmdNoheal_Click()   'if you dont want to heal, you can go right to the street or store
            MsgBox ("You decide to get out of the alley."), , ("Back to the street...")
            frmAlley.Hide
            frmStreet2.Show
            MsgBox ("Do you want to visit the old man's store again?"), , ("Store?")
            Do Until store = "yes" Or store = "no"
                store = InputBox("Type 'yes' to go to the store, or 'no' to continue...", "Yes or No?")
                If store = "yes" Then
                    frmStore2.Show
                    frmStreet2.Hide
                ElseIf store = "no" Then
                    frmStreet2.Show
                Else
                    MsgBox ("invalid entry."), , ("Error")
                End If
            Loop
    End Sub
    
    Private Sub cmdRun_Click()  'run away!  you can only run once in the game
        RunCounter = RunCounter - 1
        If RunCounter = 0 Then
            MsgBox ("You successfully ran away!"), , ("Run Away!")
            MsgBox ("What a pansy..."), , ("Pansy...")
            frmAlley.Hide
            frmStreet2.Show
        Else
            picResults.Print "You can't run anymore!  You must fight!"
        End If
        
End Sub
