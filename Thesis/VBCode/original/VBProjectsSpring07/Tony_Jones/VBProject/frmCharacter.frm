VERSION 5.00
Begin VB.Form frmCharacter 
   BackColor       =   &H00000000&
   Caption         =   "Character Selection"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBush 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   5880
      Picture         =   "frmCharacter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton cmdKen 
      Height          =   2415
      Left            =   960
      Picture         =   "frmCharacter.frx":25C3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblCharacter 
      BackColor       =   &H00000000&
      Caption         =   "Choose a Character"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program will explain the categories for the game and allow the user to see the
'character he/she picked

Private Sub cmdKen_Click()
    
    'Setting initial values of variables
    Player = 1
    Winnings = 0
    
    'Shows and hides the forms
    frmKenMoney.Show
    frmCharacter.Hide
    
    'Displays the user's name and winnings
    FName = InputBox("Please enter your first name", "First Name")
    frmKenMoney.picName.Print FName
    frmKenMoney.picWinnings.Print FormatCurrency(Winnings, 0)
    
    'Instructs the user on what to do next
    MsgBox "Click Game Board to continue to the Game Board!!!", , "Continue"
    
End Sub

Private Sub cmdBush_Click()
    
    'Setting the initial values of the variables
    Player = 2
    Winnings = 0
    
    'Shows and hides the forms
    frmBushMoney.Show
    frmCharacter.Hide
    
    'Displays the user's name and winnings
    FName = InputBox("Please enter your first name", "First Name")
    frmBushMoney.picName.Print FName
    frmBushMoney.picWinnings.Print FormatCurrency(Winnings, 0)
    
    'Instructs the user on what to do next
    MsgBox "Click Game Board to continue to the Game Board!!!", , "Continue"
    
End Sub
