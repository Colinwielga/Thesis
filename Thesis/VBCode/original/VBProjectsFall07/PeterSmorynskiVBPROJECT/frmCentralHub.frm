VERSION 5.00
Begin VB.Form frmCentralHub 
   BackColor       =   &H80000007&
   Caption         =   "POKEMON CENTRAL"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEndProgram 
      Caption         =   "End Pokemon World Simulation"
      Height          =   1335
      Left            =   4440
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "PokePhone"
      Height          =   1695
      Left            =   7320
      TabIndex        =   3
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdBattle 
      Caption         =   "Pokemon Battle"
      Height          =   1695
      Left            =   1680
      TabIndex        =   2
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdShop 
      Caption         =   "Shop at the PokeMart"
      Height          =   1695
      Left            =   7320
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdPKM 
      Caption         =   "Pokemon Selector"
      Height          =   1695
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "WELCOME TO POKEMON CENTRAL"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11055
   End
   Begin VB.Image Image4 
      Height          =   4305
      Left            =   5520
      Picture         =   "frmCentralHub.frx":0000
      Top             =   4440
      Width           =   5745
   End
   Begin VB.Image Image3 
      Height          =   5490
      Left            =   0
      Picture         =   "frmCentralHub.frx":1B2C2
      Top             =   0
      Width           =   9000
   End
   Begin VB.Image Image2 
      Height          =   4455
      Left            =   6480
      Picture         =   "frmCentralHub.frx":23950
      Top             =   -120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   6285
      Left            =   0
      Picture         =   "frmCentralHub.frx":59D5E
      Top             =   3840
      Width           =   10830
   End
End
Attribute VB_Name = "frmCentralHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPKM_Click() 'takes user to Pokemon Selector

cmdPKM.Visible = False
frmCentralHub.Hide
frmPokemonSelector.Show
MsgBox ("Some experts say the ideal Pokemon for a Pokemon Trainer should be based on that Trainer's favorite number. Let's give it a try, " & Username & "!"), , ("INSTRUCTION: SEE BELOW")
End Sub
Private Sub cmdShop_Click() 'takes user to PokeMart
cmdShop.Visible = False
frmCentralHub.Hide
frmPokeMart.Show
MsgBox (Username & " let's shop. Pokeballs catch Pokemon, Ultraballs catch them more efficiently, Potion heals damage, and Repel keeps wild Pokemon away. Oh! The owner has a message. Why don't you read it?"), , ("INSTRUCTION: SHOP")
End Sub
Private Sub cmdBattle_Click()
cmdBattle.Visible = False
frmCentralHub.Hide
frmPokemonBattle.Show
MsgBox (Username & ", " & Rivalname & " has challenged you to a one-on-one Pokemon Battle with the powerful Deoxys" & "!!! Time to Battle!"), , ("INSTRUCTION: POKEMON BATTLE")
End Sub
Private Sub cmdPhone_Click() 'takes user to PokePhone
cmdPhone.Visible = False
frmCentralHub.Hide
frmDestination.Show
MsgBox ("Time to check your PokePhone for messages. ***You have 1 NEW MESSAGE***"), , ("INSTRUCTION: POKEPHONE APPLICATION/TASK")
MsgBox ("Hi " & Username & "! I sent you a new application for your PokePhone's Map Application. Whenever you want to travel a certain distance, it'll let you know [to the nearest mile] what the first city outside that travel range is."), , ("INSTRUCTION: NEW MESSAGE: FRIEND REQUEST")
End Sub
Private Sub cmdEndProgram_Click() 'ends program
MsgBox (Username & ", I hope you've enjoyed your stay in the world of Pokemon. Simulation Terminated"), , ("See you soon!")
End
End Sub
