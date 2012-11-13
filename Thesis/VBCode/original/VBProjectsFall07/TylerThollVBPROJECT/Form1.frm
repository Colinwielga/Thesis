VERSION 5.00
Begin VB.Form frmSTART 
   BackColor       =   &H80000007&
   Caption         =   "Let's Get Started!"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0080C0FF&
      Caption         =   "Let's Begin!"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6720
      MaskColor       =   &H80000007&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label lblBach 
      BackColor       =   &H80000006&
      Caption         =   $"Form1.frx":7E87
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3015
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "frmSTART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form starts the theory session and asks for the user's name

Private Sub cmdStart_Click() 'takes the player to the Choose Menu



playername = InputBox("Please type your first name", Player) 'asks the player to type their name

    frmSTART.Hide  'moves to the choose menu
    frmChoose.Show
End Sub
