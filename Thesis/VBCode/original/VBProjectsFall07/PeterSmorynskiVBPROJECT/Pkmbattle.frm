VERSION 5.00
Begin VB.Form frmPokemonBattle 
   BackColor       =   &H8000000D&
   Caption         =   "Pokemon Battle Simulation"
   ClientHeight    =   9615
   ClientLeft      =   -60
   ClientTop       =   105
   ClientWidth     =   11220
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   11220
   Begin VB.CommandButton cmdSurrender 
      Caption         =   "Surrender"
      Height          =   855
      Left            =   8040
      TabIndex        =   2
      Top             =   3600
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      Height          =   3855
      Left            =   3000
      ScaleHeight     =   3795
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   5160
      Width           =   5775
   End
   Begin VB.CommandButton cmdAttack 
      Caption         =   "Attack"
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label lblopponent 
      BackColor       =   &H8000000D&
      Caption         =   "Your Opponent is Deoxys!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Image deoxysStrike 
      Height          =   4590
      Left            =   3840
      Picture         =   "Pkmbattle.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Image Elemental 
      Height          =   4155
      Left            =   3960
      Picture         =   "Pkmbattle.frx":BE79
      Top             =   120
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Image Deoxys 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2625
      Left            =   4320
      Picture         =   "Pkmbattle.frx":E9A7
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmPokemonBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAttack_Click() 'Determines battle results on the fly via a Do Until loop

Dim attack As Integer
Dim AttackTotal As Integer
Dim NormalAttack As Integer
Dim ElementalAttack As Integer
Dim UserLife As Integer
Dim RivalLife As Integer
Dim Damage As Integer

ElementalAttack = 2
NormalAttack = 10
UserLife = 100
RivalLife = 100
    Do Until RivalLife <= 0 Or UserLife <= 0
        attack = InputBox("Input a number to activate an attack!", "Input Number")
    If attack >= Int(50) And ElementalAttack > 0 Then
        Deoxys.Visible = False
        Elemental.Visible = True
        RivalLife = RivalLife - 40
        picResults.Print "Elemental Attack! +40 damage to Deoxys! [" & RivalLife & "/100!]"
        ElementalAttack = ElementalAttack - 1
    ElseIf attack <= Int(50) Then
        Deoxys.Visible = False
        deoxysStrike.Visible = True
        RivalLife = RivalLife - 10
        picResults.Print "Normal Attack! +10 damage to Deoxys! [" & RivalLife & "/100!]"
        NormalAttack = NormalAttack - 1
        MsgBox ("Enemy Pokemon struck back!"), , ("Counterattack!")
        UserLife = UserLife - 40
        picResults.Print "Elemental Attack! +40 damage to your Pokemon! [" & UserLife & "/100!]"
    ElseIf ElementalAttack <= 0 Then
        MsgBox ("Your Pokemon has exhausted its elemental powers. It used a Normal Attack instead"), , ("Your Pokemon was only able to do a Normal Attack!")
        Deoxys.Visible = False
        deoxysStrike.Visible = True
        RivalLife = RivalLife - 10
        picResults.Print "Normal Attack! +10 damage to Deoxys! [" & RivalLife & "/100!]"
        NormalAttack = NormalAttack - 1
        MsgBox ("Enemy Pokemon struck back!"), , ("Counterattack!")
        UserLife = UserLife - 15
        picResults.Print "Normal Attack! +15 damage to your Pokemon! [", UserLife & "/100!]"
    End If
    Loop
If RivalLife And UserLife <= 0 Then 'ending conditions
    MsgBox ("Both Pokemon fainted. It's a draw!"), , ("THE BATTLE IS OVER!")
    MsgBox ("A lucky break, kid. Next match won't be a draw!"), , (Rivalname & " says")
    ElseIf RivalLife <= 0 Then
    MsgBox (Rivalname & "'s Deoxys has fainted." & Username & " is victorious!"), , ("THE BATTLE IS OVER!")
    MsgBox ("A lucky break, kid! Next time I won't hold back!"), , (Rivalname & " says")
    ElseIf UserLife <= 0 Then
    MsgBox (Username & "'s Pokemon has fainted." & Rivalname & " is victorious!"), , ("THE BATTLE IS OVER!")
    MsgBox ("Nice try, kid! I can tell you've been training hard. I look forward to another match."), , (Rivalname & " says")
    MsgBox ("An impressive battle," & Username & "Better luck next time!"), , ("INSTRUCTION: CHOOSE YOUR NEXT DESTINATION")
End If
    frmPokemonBattle.Hide
    frmCentralHub.Show
MsgBox ("Where to next?"), , ("INSTRUCTION: CHOOSE YOUR NEXT DESTINATION")
End Sub
Private Sub cmdSurrender_Click() 'return to pokemon central
    MsgBox ("Running away?! Challenge me again when you're up to it, kid!"), , (Rivalname & " says")
frmPokemonBattle.Hide
frmCentralHub.Show
MsgBox ("An impressive battle," & Username & "I think you'll win next time!"), , ("INSTRUCTION: CHOOSE YOUR NEXT DESTINATION")
End Sub

