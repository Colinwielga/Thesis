VERSION 5.00
Begin VB.Form frmDisneyCastle 
   Caption         =   "Welcome to Disney Castle"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   Picture         =   "frmDisneyCastle.frx":0000
   ScaleHeight     =   6150
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMovies 
      BackColor       =   &H0080FF80&
      Caption         =   "Search for Disney Movies"
      Height          =   1095
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Leave Disney Castle"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H00FF8080&
      Caption         =   "Buy Souvienirs"
      Height          =   1095
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdVillians 
      BackColor       =   &H008080FF&
      Caption         =   "Meet Villains"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdGame 
      BackColor       =   &H00FFFF80&
      Caption         =   "Play a Game"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrincesses 
      BackColor       =   &H0080FFFF&
      Caption         =   "Meet Disney Princesses"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmDisneyCastle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Project
'Disney Castle
'Lori Nohner
'Written March 17, 2008
'Objective- allows user to access different forms in the project
Option Explicit

Private Sub cmdBuy_Click()
    frmDisneyCastle.Hide 'hides Disney Castle page
    frmSouviner.Show 'goes to souviner page
    Money = InputBox("How much money do you want to spend?", "Enter Dollar Amount")
       
End Sub

Private Sub cmdExit_Click()
    End ' ends program
End Sub

Private Sub cmdGame_Click()
    frmDisneyCastle.Hide 'hides Disney Castle page
    frmGame.Show ' goes to hereos page
End Sub


Private Sub cmdMovies_Click()
    frmDisneyCastle.Hide 'hides Disney Castle page
    frmMovies.Show 'goes to movie page
    MsgBox "Here you can find out what classic Disney movies you are missing.", , "Disney Movies"
End Sub

Private Sub cmdPrincesses_Click()
    frmDisneyCastle.Hide 'hides Disney Castle page
    frmPrincesses.Show ' goes to princess page
    
End Sub

Private Sub cmdVillians_Click()
    frmDisneyCastle.Hide 'hides Disney Castle page
    frmVillains.Show 'goes to villains page
End Sub


Private Sub Form_Load()

End Sub
