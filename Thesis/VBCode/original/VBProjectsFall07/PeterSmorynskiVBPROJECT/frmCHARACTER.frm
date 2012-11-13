VERSION 5.00
Begin VB.Form frmCHARACTER 
   BackColor       =   &H8000000D&
   Caption         =   "CHARACTER SELECTION"
   ClientHeight    =   9600
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   12525
   LinkTopic       =   "Form2"
   ScaleHeight     =   9600
   ScaleWidth      =   12525
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue Simulation"
      Height          =   735
      Left            =   5160
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Game"
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   8520
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picResults2 
      Height          =   735
      Left            =   5760
      ScaleHeight     =   675
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdFemale 
      Caption         =   "Trainer (Female)"
      Height          =   1455
      Left            =   7560
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdMale 
      BackColor       =   &H80000001&
      Caption         =   "Trainer (Male)"
      Height          =   1455
      Left            =   4920
      MaskColor       =   &H00400000&
      TabIndex        =   0
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Image Image2Result 
      Height          =   3660
      Left            =   6120
      Picture         =   "frmCHARACTER.frx":0000
      Top             =   1200
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Image Image1Result 
      Height          =   3660
      Left            =   6120
      Picture         =   "frmCHARACTER.frx":1F1E
      Top             =   1200
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image2female 
      Height          =   3660
      Left            =   7560
      Picture         =   "frmCHARACTER.frx":3D20
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Image Image1male 
      Height          =   3660
      Left            =   4920
      Picture         =   "frmCHARACTER.frx":5C3E
      Top             =   1200
      Width           =   1635
   End
End
Attribute VB_Name = "frmCHARACTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdFemale_Click() 'Brings up female image, asks for name and rival name.
cmdFemale.Visible = False
cmdMale.Visible = False
Image2female.Visible = False
Image1male.Visible = False
Username = InputBox("Enter your Female Trainer's name", "Input a Name")
Image2Result.Visible = True
picResults2.Print Username
Rivalname = InputBox("What is your RIVAL's name?", "Input a Name")
cmdContinue.Visible = True
cmdQuit.Visible = True
End Sub
Private Sub cmdMale_Click() ''Brings up male image, asks for name and rival name.
cmdMale.Visible = False
cmdFemale.Visible = False
Image2female.Visible = False
Image1male.Visible = False
Username = InputBox("Enter your Male Trainer's name", "Input a Name")
Image1Result.Visible = True
picResults2.Print Username
Rivalname = InputBox("What is your RIVAL's name?", "Input a Name")
cmdContinue.Visible = True
cmdQuit.Visible = True
End Sub
Private Sub cmdContinue_Click() 'Go to Pokemon Central
frmCHARACTER.Hide
frmCentralHub.Show
MsgBox ("Welcome to Pokemon Central, " & Username & "!" & "This in the main hub for all simulations relating to the life of a Pokemon Trainer. You will only be able to experience each simulation once per Program Start, so take your time. You're free to go in any order you like! Have fun exploring!"), , ("INSTRUCTION: CHOOSE YOUR DESTINATION!")
End Sub
Private Sub cmdQuit_Click() 'end program
End
End Sub

