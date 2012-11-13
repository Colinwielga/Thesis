VERSION 5.00
Begin VB.Form frmKnightStart 
   BackColor       =   &H00000000&
   Caption         =   "Knight's Tale"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmdDie2 
      Caption         =   "Deny Mission. Don't Save the Princess."
      Height          =   735
      Left            =   4920
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdSavePrincess2 
      Caption         =   "Accept the Mission! Save the Princess!"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdDie1 
      Caption         =   "Don't Save the Princess. Live with Toads and Warts Forever!"
      Height          =   975
      Left            =   4920
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdSavePrincess1 
      Caption         =   "Go Save the Princess! Break the Curse!"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picWizardQuote 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      ScaleHeight     =   1275
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   1800
      Width           =   5055
   End
   Begin VB.CommandButton cmdnotalk 
      Caption         =   "Don't Talk to the Wizard And Continue"
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdTalk 
      Caption         =   "Talk to The Wizard"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape shpquote1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   5310
      Left            =   600
      Picture         =   "frmKnightStart.frx":0000
      Top             =   2880
      Width           =   3435
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
      Caption         =   "One a dreadful day of October, the Wizard of Mystic Forest Came to You. "
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmKnightStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Katie Deeney & Elise Generex
    'Create a Story
    'Date Done: 10/10/2009
    'This form the knight faces the wizard
    'The wizard will send the knight off on a mission
    'or the wizard will cast a spell on you
    
Private Sub cmdDie1_Click()
    'This button means that you died and have to start over
    MsgBox "Due to your warts, no one likes you anymore.  You are no longer noble. Due to these circumstances Sir " & CharacterName & ", you become depressed and live with the Toads Forever!", , "Bad Choice"
    MsgBox "This is where your story ends. Start Over", , "Sotry Ends"
    frmKnightStart.Hide
    cmdTalk.Enabled = True
    picWizardQuote.Cls
    frmWelcome.Show
    cmdDie1.Visible = False
    cmdSavePrincess1.Visible = False
    
    
End Sub

Private Sub cmdDie2_Click()
    'This button means you died and have to start over
    MsgBox "All of Your Friends Believe you are a Coward, You get Depressed and Commit Suicide.", , "Bad Choice"
    MsgBox "This is where your story ends. Start Over", , "Sotry Ends"
    cmdnotalk.Enabled = True
    picWizardQuote.Cls
    frmKnightStart.Hide
    frmWelcome.Show
    cmdDie2.Visible = False
    cmdSavePrincess2.Visible = False
End Sub

Private Sub cmdnotalk_Click()
'You didn't want to talk to the wizard
'Because of this, this button cast a spell on the user
 picWizardQuote.Print "A curse shall be thrashed upon thee on this very day!"
 picWizardQuote.Print "You shall be covered in warts!"
 picWizardQuote.Print "The only way to break the Curse is to save the captured "
 picWizardQuote.Print "Princess from the dragon!"
 cmdDie1.Visible = True
 cmdSavePrincess1.Visible = True
 cmdTalk.Enabled = False
 cmdDie2.Visible = False
 cmdSavePrincess2.Visible = False
 
End Sub

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub cmdSavePrincess1_Click()
 cmdTalk.Enabled = True
 cmdDie1.Visible = False
 cmdSavePrincess1.Visible = False
 frmKnightStart.Hide
 frmSavePrincess1.Show
 picWizardQuote.Cls
End Sub

Private Sub cmdSavePrincess2_Click()
    cmdnotalk.Enabled = True
    cmdDie2.Visible = False
    cmdSavePrincess2.Visible = False
    frmKnightStart.Hide
    frmSavePrincess2.Show
    picWizardQuote.Cls
    
End Sub

Private Sub cmdTalk_Click()
'This button sends the user off on a misson.
 picWizardQuote.Print "Thy Fair Maiden is Trapped Deep in a Cave!"
 picWizardQuote.Print "Thoust shall Go on a Mission To Save Her."
 picWizardQuote.Print "Beware! Thoust shall encounter Dangers!"
 cmdnotalk.Enabled = False
 cmdDie2.Visible = True
 cmdSavePrincess2.Visible = True
 cmdDie1.Visible = False
 cmdSavePrincess1.Visible = False
End Sub

