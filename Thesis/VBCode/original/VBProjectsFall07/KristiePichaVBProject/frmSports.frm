VERSION 5.00
Begin VB.Form frmSports 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Back to Menu"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdManager 
      BackColor       =   &H00FF0000&
      Caption         =   "Football Managing"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCheerleading 
      BackColor       =   &H00FF0000&
      Caption         =   "Cheerleading"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdBasketball 
      BackColor       =   &H000000FF&
      Caption         =   "Basketball"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdVolleyball 
      BackColor       =   &H000000FF&
      Caption         =   "Volleyball"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Caption         =   "Sports that Kristie was involved in during high school."
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmSports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBasketball_Click()
MsgBox "Kristie played Basketball for four years in 7th-10th grades., Basketball"
End Sub

Private Sub cmdCheerleading_Click()
MsgBox "Kristie was a cheerleader for one year in 8th grade., Cheerleading"
End Sub

Private Sub cmdManager_Click()
MsgBox "Kristie was a football manager for one year in 9th grade., Football Manager"
End Sub

Private Sub cmdQuit_Click()
frmSports.Hide
frmMenu.Show
End Sub

Private Sub cmdVolleyball_Click()
MsgBox "Kristie played Volleyball for one year in 7th grade."
End Sub


Private Sub Form_Load()

End Sub
