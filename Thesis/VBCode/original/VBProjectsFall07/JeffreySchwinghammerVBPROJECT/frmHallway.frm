VERSION 5.00
Begin VB.Form frmHallway 
   BackColor       =   &H80000007&
   Caption         =   "Hallway"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   2055
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   8550
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   360
      Picture         =   "frmHallway.frx":0000
      ScaleHeight     =   7335
      ScaleWidth      =   7695
      TabIndex        =   3
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue On"
      Height          =   735
      Left            =   6600
      TabIndex        =   2
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Go Back to Previous Room"
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   7680
      Width           =   1455
   End
   Begin VB.PictureBox picHalltxt 
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   8520
      Width           =   7575
   End
End
Attribute VB_Name = "frmHallway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdContinue_Click()
frmLab.Show
frmHallway.Hide
End Sub

Private Sub cmdReturn_Click()
frmHub.Show
frmHallway.Hide
End Sub

Private Sub Form_activate()
    picHalltxt.Cls
    picHalltxt.Print "You are taken aback, as you step into an odd hallway. It looks"
    picHalltxt.Print "it is apart of an entirely different building. Maybe apart of"
    picHalltxt.Print "scientist's lab or a military research facility."
End Sub
