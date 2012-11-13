VERSION 5.00
Begin VB.Form frmDuckBill 
   BackColor       =   &H00000040&
   Caption         =   "Did that really just happen?"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   Picture         =   "frmDuckBill.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00004080&
      Caption         =   "Find the Time Machine"
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdExplore 
      BackColor       =   &H00004080&
      Caption         =   "Explore"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label lblDinosaur 
      BackColor       =   &H00004080&
      Caption         =   $"frmDuckBill.frx":BD02
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   4095
   End
End
Attribute VB_Name = "frmDuckBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExplore_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmDuckBill.Visible = False
    frmMountain.Visible = True
    
    'Message boxes about journey
    MsgBox "Whee.... " & YourName & ", walking through the jungle is fun.", , "Walking"
    MsgBox "Hey! There's a mountain in the distance. Let's go there!", , "Mountain!"
End Sub

Private Sub cmdReturn_Click()
    'Conceals one form to reveal another, based on the button pushed.
    frmDuckBill.Visible = False
    frmTimeMachine.Visible = True
    
End Sub
