VERSION 5.00
Begin VB.Form frmFightSword2 
   BackColor       =   &H0000FFFF&
   Caption         =   "Fight with the Sword!"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdbazillion 
      Caption         =   "A Bizillion Kajillion Times"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "Eight Times"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdOnce 
      Caption         =   "Once"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "How many times would you like to stab the dragon?"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   1560
      Index           =   0
      Left            =   3240
      Picture         =   "frmFightSword2.frx":0000
      Top             =   2880
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   3120
      Picture         =   "frmFightSword2.frx":0A94
      Top             =   1080
      Width           =   4500
   End
End
Attribute VB_Name = "frmFightSword2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'This is a form where the user has to fight the dragon with
'a sword. The user has 3 buitton options. If the user
'Picks the first one, the user dies
'If the user picks the second one, the user kills the dragon
'If the user picks the third one, he/she kilss the dragon but
'dies in the process of doing so
'The user will then return to the beginning and can try again
Private Sub cmdbazillion_Click()
    MsgBox "Wow! You killed the dragon in the first ten stabs, but you kept going until a bazillion kajillion, so unfortunately you died of exhaustion. However, the princess is saved.", , "Bad Choice"
    MsgBox "This is where your story ends, Start over", , "Story Ends"
    Inventory = ""
    frmFightSword2.Hide
    frmWelcome.Show
End Sub

Private Sub cmdeight_Click()
    MsgBox "Nice Work! The Dragon has been slain! You saved the Princess!", , "Whoot!"
    MsgBox "This is where your story happily ends, Start over", , "Story Ends"
    Inventory = ""
    frmFightSword.Hide
    frmWelcome.Show
End Sub

Private Sub cmdOnce_Click()
    MsgBox "You're strong, but not that strong.  You didn't slay the dragon so the dragon slayed you. Many attended your funeral.", , "Bad Choice"
    MsgBox "This is where your story ends, Start over", , "Story Ends"
    Inventory = ""
    frmFightSword2.Hide
    frmWelcome.Show
End Sub

Private Sub CmdQuit_Click()
    End
End Sub
