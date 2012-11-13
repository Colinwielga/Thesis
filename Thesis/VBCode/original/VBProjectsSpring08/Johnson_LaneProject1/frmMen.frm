VERSION 5.00
Begin VB.Form frmMen 
   Caption         =   "Men's  Section"
   ClientHeight    =   8940
   ClientLeft      =   1125
   ClientTop       =   1620
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   Picture         =   "frmMen.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11970
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   1935
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMen.frx":36008
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccessories 
      Caption         =   "Accessories"
      Height          =   1815
      Left            =   8040
      Picture         =   "frmMen.frx":3643E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdShoes 
      Caption         =   "Shoes"
      Height          =   1815
      Left            =   5400
      Picture         =   "frmMen.frx":36A8A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdApparell 
      Caption         =   "Apparel"
      Height          =   1815
      Left            =   2520
      Picture         =   "frmMen.frx":37551
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmMen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmMen
'Author: Sean Johnson and Nick Lane
'Date Written: Friday March 14th, 2007
'Objective of form: this form is the entry point to the men's section of the store.
'                   it allows the user to explore different sections of the store
'                   relating to men. it allows the users to select either the Apparel,
'                   shoes, or accessories for men.


Private Sub cmdAccessories_Click()
    'hides this form and takes the user to the accessories form
    frmMenAccessories.Show
    frmMen.Hide
End Sub

Private Sub cmdApparell_Click()
    'hides this form and takes the user to the accessories form
    frmMenApparel.Show
    frmMen.Hide
End Sub

Private Sub cmdBack_Click()
    'allows user to return to previous form
    frmMen.Hide
    frmSecondPage.Show
End Sub

Private Sub cmdShoes_Click()
    'hides this form and takes the user to the accessories form
    frmShoes.Show
    frmMen.Hide
End Sub


