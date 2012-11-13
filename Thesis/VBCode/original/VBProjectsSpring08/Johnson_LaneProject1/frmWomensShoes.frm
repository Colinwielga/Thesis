VERSION 5.00
Begin VB.Form frmWomensShoes 
   Caption         =   "Women's Shoes"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   Picture         =   "frmWomensShoes.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000013&
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdBasketball 
      BackColor       =   &H80000013&
      Caption         =   "Basketball"
      Height          =   1215
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdRunning 
      BackColor       =   &H80000013&
      Caption         =   "Running"
      Height          =   1215
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdCasual 
      BackColor       =   &H80000013&
      Caption         =   "Casual"
      Height          =   1215
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "frmWomensShoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmWomensShoes
'Author: Sean Johnson and Nick Lane
'Date Written: Friday March 14th, 2007
'Objective of form: this particular form allows the user to enter into the Women Shoe's section of the store.
'                   from this form, the user can decide and select their desired type of shoes they wish to explore.

Private Sub cmdBack_Click()
'hides this form and displays the previous form
frmWomensShoes.Hide
frmWomen.Show
End Sub

Private Sub cmdBasketball_Click()
'hides this form and displays the women basketball shoes form
frmWomensShoes.Hide
frmWomensBasketball.Show
End Sub

Private Sub cmdCasual_Click()
'hides this form and displays the women casual shoes form
frmWomensShoes.Hide
frmWomensCasual.Show
End Sub

Private Sub cmdRunning_Click()
'hides this form and displays the women casual shoes form
frmWomensShoes.Hide
frmWomensRunning.Show
End Sub
