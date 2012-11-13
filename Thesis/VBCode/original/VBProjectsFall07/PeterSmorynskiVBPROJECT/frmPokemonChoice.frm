VERSION 5.00
Begin VB.Form frmPokemonSelector 
   BackColor       =   &H8000000D&
   Caption         =   "Pokemon Selector"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRtrHub 
      Caption         =   "Return to Pokemon Central"
      Height          =   1095
      Left            =   4560
      TabIndex        =   3
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox txtPokemonNumber 
      Height          =   1695
      Left            =   3960
      TabIndex        =   2
      Top             =   2400
      Width           =   3495
   End
   Begin VB.CommandButton cmdPKMpick 
      BackColor       =   &H80000009&
      Caption         =   "Enter a Number Above and Click Here for your Ideal Pokemon!"
      Height          =   1095
      Left            =   4080
      MaskColor       =   &H8000000D&
      TabIndex        =   1
      Top             =   4440
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   3960
      ScaleHeight     =   915
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Image Bulbasaurpic 
      Height          =   1410
      Left            =   8160
      Picture         =   "frmPokemonChoice.frx":0000
      Top             =   2520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Squirtlepic 
      Height          =   1545
      Left            =   6360
      Picture         =   "frmPokemonChoice.frx":1A11
      Top             =   2400
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Image Pikachupic 
      Height          =   1575
      Left            =   4560
      Picture         =   "frmPokemonChoice.frx":2975
      Top             =   2400
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Image Charmanderpic 
      Height          =   1500
      Left            =   3000
      Picture         =   "frmPokemonChoice.frx":DE2F
      Top             =   2400
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "frmPokemonSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtPokemonNumber_Change()
Dim Pokemon As Integer
End Sub
Private Sub cmdPKMpick_Click() 'Takes textbox number and searches for a match via Select Case
Dim Pokemon As Single
Pokemon = txtPokemonNumber.Text
    cmdPKMpick.Visible = False
    txtPokemonNumber.Visible = False
    Select Case Pokemon
        
        Case 50 To 100
            picResults.Print "CHARMANDER is a perfect fit. You have a"
            picResults.Print "firey spirit, don't you?"
            Charmanderpic.Visible = True
        Case 25 To 50
            picResults.Print "SQUIRTLE works for you."
            picResults.Print "You must enjoy water sports."
            Squirtlepic.Visible = True
        Case 0 To 25
            picResults.Print "The electrifying PIKACHU is yours."
            picResults.Print "You probably have an odd but"
            picResults.Print "electric personality."
            Pikachupic.Visible = True
        Case Else
            picResults.Print "BULBASAUR, a great match. You're lazy"
            picResults.Print "but reliable."
            Bulbasaurpic.Visible = True
    End Select
End Sub

Private Sub cmdRtrHub_Click() ' Return to pokemon central
frmPokemonSelector.Hide
frmCentralHub.Show
Dim Opinion
MsgBox ("Welcome back to Pokemon Central! Did you like the Pokemon the simulator chose for you?"), , ("INSTRUCTION: QUESTION FOR USER")
Opinion = InputBox("1 for YES, 2 for NO", "Question: Did you like the Pokemon chosen for you?")
    If Opinion = 1 Then
    MsgBox ("Yes, I think you and your Pokemon look quite alike!"), , ("SIMULATOR REPLY")
    Else
    MsgBox ("I'm not very fond of it, either. Luckily this is just a simulation."), , ("SIMULATOR REPLY")
    
End If
End Sub
