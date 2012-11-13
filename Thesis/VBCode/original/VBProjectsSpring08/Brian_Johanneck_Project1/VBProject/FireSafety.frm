VERSION 5.00
Begin VB.Form FireSafety 
   BackColor       =   &H000000FF&
   Caption         =   "Fire Safety"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   Picture         =   "FireSafety.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   4560
      Picture         =   "FireSafety.frx":B9AF6
      ScaleHeight     =   4995
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   3840
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "What you should do if there is a fire in your house."
      Height          =   1575
      Left            =   8280
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tips for what to do to prepair in case of a fire."
      Height          =   1575
      Left            =   5160
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Back 
      Caption         =   "Go  back to main menu."
      Height          =   1455
      Left            =   6720
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "FireSafety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Back_Click()
FireSafety.Hide
Form1.Show
End Sub

Private Sub Command1_Click()
MsgBox ("Click ok when your ready for the next tip.")
MsgBox ("Have an escape plan that you practice every month.")
MsgBox ("Have a working fire extinguisher.")
MsgBox ("Have a smoke detector with working a workin battery.")
End Sub

Private Sub Command2_Click()
MsgBox ("Click ok when your ready for the next tip.")
MsgBox ("Use your escape plan and get out of the house.")
MsgBox ("When you think there is a fire in your house make sure to get out right away.")
MsgBox ("Touch door handles so you know if there may be fire in the next room before you open the door.")
MsgBox ("Meet at a designated place outside the house so everyones know everybody is out.")
MsgBox ("Do not go back in!")
End Sub
