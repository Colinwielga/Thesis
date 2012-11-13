VERSION 5.00
Begin VB.Form frmQuest 
   Caption         =   "Quest"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Picture         =   "frmQuest.frx":0000
   ScaleHeight     =   6090
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEnemyHealth 
      Height          =   255
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdAttack 
      Caption         =   "Attack!"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start!"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.PictureBox picMyHealth 
      Height          =   255
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdAttributes 
      Caption         =   "Character Attributes and Details"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblEnemyHealth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enemy Health:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblMyHealth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Health:"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   5400
      Width           =   855
   End
   Begin VB.Image imgRat 
      BorderStyle     =   1  'Fixed Single
      Height          =   2010
      Left            =   3240
      Picture         =   "frmQuest.frx":CF8C
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   8520
      Y1              =   5280
      Y2              =   5280
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmQuest
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of form: This form brings the user into a battle where they can gain strength.
        'The user has the advantage, which helps them so they don't lose all of their cash if they lose.

Option Explicit

Private Sub cmdAttack_Click()
    Dim MyDmg As Integer, EnemyDmg As Integer   'declares the variables MyDmg and EnemyDmg
    Dim Range1 As Integer, Range2 As Integer    'declares the variables range for damage.
    
    CTR = 0 'Sets the counter equal to zero.
    
    Do Until CTR = 1
    
        picMyHealth.Cls 'Clears/resets my health for the next count.
        picEnemyHealth.Cls  'Clears/resets the enemy health for the next count.
        
        Strength = Strength + 1 'Adds a strength point everytime the user attacks.
        MsgBox "You gained 1 strength point!", , "+1 Strength!" 'Notifies the user that they gained an extra strength point.
        
        CTR = CTR + 1   'Increments counter.
        
        Randomize (Range1)  'Randomizes the Range1 variable.
        Range1 = CInt(Int((30 * Rnd()) + 1))    'Sets the range for Range1 to 1 - 30 and is randomized.
        
        Randomize (Range2)  'Randomizes the Range2 variable.
        Range2 = CInt(Int((40 * Rnd()) + 1))    'Sets the range for Range1 to 1 - 40 and is randomized.
        
        MyDmg = Range1  'Assigns Range1 to MyDmg.
        EnemyDmg = Range2   'Assings Range2 to EnemyDmg.
        
        MyHealth = MyHealth - MyDmg 'Computes the amount of user's health left after each attack.
        EnemyHealth = EnemyHealth - EnemyDmg    'Computes the amount of enemy health left after each attack.
        
        Select Case MyHealth
                Case Is <= 0    'If the user's health is less than or equal to zero, then they lose.
                    MsgBox "You lost! All of your current cash has been taken away! You can head back to town...", , "Defeat!"  'Notifies the user of a loss.
                    cmdAttack.Enabled = False   'Disables the attack button once a loss occurs.
                    Cash = 0    'If the user loses, then they lose all of their cash.
                Case Is < 100   'If the user's health is less than 100, then they will be notified of the damage done by each fighter.
                    MsgBox "The enemy did " & MyDmg & " damage to you.", , "Damage Report"
                    picMyHealth.Print MyHealth  'Prints user health after every attack.
        End Select
    
        Select Case EnemyHealth
                Case Is <= 0    'If the enemy's health is less than or equal to zero, then the user wins.
                    MsgBox "Congratulations, you won! You can head back to town!", , "Success!"
                    cmdAttack.Enabled = False
                    imgRat.Visible = False
                Case Is < 100
                    MsgBox "You did " & EnemyDmg & " damage to the enemy.", , "Damage Report"
                    picEnemyHealth.Print EnemyHealth    'Prints enemy health after every attack.
        End Select
        
    Loop
    
End Sub

Private Sub cmdAttributes_Click()

    frmAttributes.Show
    
End Sub

Private Sub cmdCharacter_Click()

    frmCharacter.Show
    
End Sub

Private Sub cmdBack_Click()
    frmQuest.Hide
    
    If MyHealth < 100 Then
        MsgBox "Your health is below 100, you should visit the hospital!", , "Low Health!"  'If the user's health is below 100, then they are told to visit the hospital.
    End If
    
End Sub

Private Sub cmdStart_Click()

    MsgBox "A rat is attacking you!", , "You're Under Attack!" 'Warns the user of attack.
    imgRat.Visible = True
    
End Sub

Private Sub Form_Load()

    MyHealth = 100  'Sets the user's health to 100.
    EnemyHealth = 100   'Sets the enemy's health to 100.
    
End Sub
