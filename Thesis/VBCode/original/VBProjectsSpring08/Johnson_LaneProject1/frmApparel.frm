VERSION 5.00
Begin VB.Form frmMenApparel 
   Caption         =   "Men's Apparel"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   Picture         =   "frmApparel.frx":0000
   ScaleHeight     =   8640
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBMen 
      BackColor       =   &H80000013&
      Caption         =   "Return to Men's Page"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdShorts 
      BackColor       =   &H80000013&
      Caption         =   "Short"
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSweats 
      BackColor       =   &H80000013&
      Caption         =   "Sweats"
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdShirt 
      BackColor       =   &H80000013&
      Caption         =   "Shirts"
      Height          =   975
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdJerseys 
      BackColor       =   &H80000013&
      Caption         =   "Jersey"
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdHat 
      BackColor       =   &H80000013&
      Caption         =   "Hat"
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1215
   End
End
Attribute VB_Name = "frmMenApparel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmMenApparel
'Author: Sean Johnson and Nick Lane
'Date Written: Friday March 14th, 2007
'Objective of form: this particular form allows the user to enter into the Men's Apparel section of the store.
'                   from this form, the user can decide and select their desired type of apparel they which to buy.

Private Sub cmdBMen_Click()
'allows the user to return to previous form
frmMen.Show
frmMenApparel.Hide
End Sub

Private Sub cmdHat_Click()
'hides this form and takes the user to the Hat form
frmMenApparel.Hide
frmHat.Show
End Sub

Private Sub cmdJerseys_Click()
'hides this form and takes the user to the Jersey form
frmMenApparel.Hide
frmJersey.Show
End Sub

Private Sub cmdShirt_Click()
'hides this form and takes the user to the Shirt form
frmMenApparel.Hide
frmShirt.Show
End Sub

Private Sub cmdShorts_Click()
'hides this form and takes the user to the Shorts form
frmMenApparel.Hide
frmShorts.Show
End Sub

Private Sub cmdSweats_Click()
'hides this form and takes the user to the Sweats form
frmMenApparel.Hide
frmSweats.Show
End Sub
