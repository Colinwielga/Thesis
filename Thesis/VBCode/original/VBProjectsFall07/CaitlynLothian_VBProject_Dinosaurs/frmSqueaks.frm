VERSION 5.00
Begin VB.Form frmSqueaks 
   Caption         =   "Squeakers"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmSqueaks.frx":0000
   ScaleHeight     =   7320
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdThrow 
      BackColor       =   &H00404080&
      Caption         =   "Throw a stick!"
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdFreeze 
      BackColor       =   &H00404080&
      Caption         =   "Freeze!"
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label lblSqueaks 
      BackColor       =   &H8000000E&
      Caption         =   $"frmSqueaks.frx":1F285
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   4695
   End
End
Attribute VB_Name = "frmSqueaks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFreeze_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmSqueaks.Visible = False
    frmFreeze.Visible = True
    
    'Message boxes about actions to distract baby dinosaurs
    MsgBox "You freeze in your spot, " & YourName & ", and they continue playing. They haven't seen you. You breathe a sigh of relief and relax slightly.", , "Safe?"
    MsgBox "The babies might not have seen you, but their mother did, and she's very upset that you're in her territory.", , "Mom's are always worse."
End Sub

Private Sub cmdThrow_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmSqueaks.Visible = False
    frmChasm.Visible = True
    
    'Message boxes about actions to distract baby dinosaurs
    MsgBox "You pick up a large stick next to you, and toss it so it sails over their heads and lands in the brush behind them. They stop playing, and head towards that bush to see what made the noise.", , "Fetch"
    MsgBox "It worked, " & YourName & "! Run for it!", , "Run Away!"
    
End Sub


