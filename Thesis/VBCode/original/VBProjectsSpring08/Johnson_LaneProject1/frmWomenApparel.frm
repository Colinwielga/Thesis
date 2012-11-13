VERSION 5.00
Begin VB.Form frmWomenApparel 
   Caption         =   "Women's Apparel"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   Picture         =   "frmWomenApparel.frx":0000
   ScaleHeight     =   9105
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBWomen 
      BackColor       =   &H0080FFFF&
      Caption         =   "Back to Womens Page"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdHeadwear 
      BackColor       =   &H0080FFFF&
      Caption         =   "Headwear"
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmdSDresses 
      BackColor       =   &H0080FFFF&
      Caption         =   "Skirt/ Dresses"
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdWarmUp 
      BackColor       =   &H0080FFFF&
      Caption         =   "Shirts"
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton frmCapris 
      BackColor       =   &H0080FFFF&
      Caption         =   "Capris"
      Height          =   915
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   2115
   End
   Begin VB.CommandButton cmdSJackets 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sweater/ Jackets"
      Height          =   975
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "frmWomenApparel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmMenApparel
'Author: Sean Johnson and Nick Lane
'Date Written: Friday March 14th, 2007
'Objective of form: this particular form allows the user to enter into the women Apparel section of the store.
'                   from this form, the user can decide and select their desired type of apparel they which to buy.


Private Sub cmdBWomen_Click()
'hides this form and displays the previous form
frmWomen.Show
frmWomenApparel.Hide
End Sub

Private Sub cmdHeadwear_Click()
'hides this form and displays the women headwear form
frmWomenApparel.Hide
frmHeadwear.Show
End Sub

Private Sub cmdSDresses_Click()
'hides this form and displays the women dresses form
frmWomenApparel.Hide
frmdresses.Show
End Sub

Private Sub cmdSJackets_Click()
'hides this form and displays the women sweaters form
frmWomenApparel.Hide
frmSweaters.Show
End Sub

Private Sub cmdWarmUp_Click()
'hides this form and displays the women shirt form
frmWomenApparel.Hide
frmShirts.Show
End Sub

Private Sub frmCapris_Click()
'hides this form and displays the women capris form
frmWomenApparel.Hide
frmCap.Show
End Sub
