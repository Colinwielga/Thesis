VERSION 5.00
Begin VB.Form frmShoes 
   Caption         =   "Shoe Main"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   Picture         =   "frmShoes.frx":0000
   ScaleHeight     =   8175
   ScaleWidth      =   12990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBmen 
      BackColor       =   &H80000013&
      Caption         =   "Back"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdCasual 
      BackColor       =   &H80000013&
      Caption         =   "Casual"
      Height          =   1455
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmdRunning 
      BackColor       =   &H80000013&
      Caption         =   "Running"
      Height          =   1455
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdBasketball 
      BackColor       =   &H80000013&
      Caption         =   "Basketball"
      Height          =   1455
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmShoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmShoes
'Author: Sean Johnson and Nick Lane
'Date Written: Friday March 14th, 2007
'Objective of form: this particular form allows the user to enter into the Men's Shoe's section of the store.
'                   from this form, the user can decide and select their desired type of shoes they wish to explore

Private Sub cmdBasketball_Click()
'hides this form and displays the basketball shoes form
frmBshoe.Show
frmShoes.Hide
End Sub


Private Sub cmdBMen_Click()
'hides this form and displays the previous form
frmMen.Show
frmShoes.Hide
End Sub

Private Sub cmdCasual_Click()
'hides this form and displays the casual shoes form
frmCshoes.Show
frmShoes.Hide
End Sub

Private Sub cmdRunning_Click()
'hides this form and displays the running shoes form
frmRshoes.Show
frmShoes.Hide
End Sub

