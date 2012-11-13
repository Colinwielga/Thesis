VERSION 5.00
Begin VB.Form frmWomen 
   Caption         =   "Women's Section"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   Picture         =   "frmWomen.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to Previous Page"
      Height          =   615
      Left            =   9600
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdAccessories 
      Caption         =   "Accessories"
      Height          =   1695
      Left            =   840
      Picture         =   "frmWomen.frx":2AA84E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdShoes 
      Caption         =   "Shoes"
      Height          =   1695
      Left            =   2280
      Picture         =   "frmWomen.frx":2AB544
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdApparell 
      Caption         =   "Apparel"
      Height          =   1695
      Left            =   840
      Picture         =   "frmWomen.frx":2AC126
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmWomen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Nike Town
'Form name: frmWomen
'Author: Sean Johnson and Nick Lane
'Date Written: Friday March 14th, 2007
'Objective of form: this form is the entry point to the women section of the store.
'                   it allows the user to explore different sections of the store
'                   relating to women. it allows the users to select either the Apparel,
'                   shoes, or accessories for women.

Private Sub cmdAccessories_Click()

    'hides this form and displays the women accessories form
    frmWomenAccessories.Show
    frmWomen.Hide
End Sub

Private Sub cmdApparell_Click()

    'hides this form and displays the women apparel form
    frmWomenApparel.Show
    frmWomen.Hide
End Sub

Private Sub cmdBack_Click()

    'hide this form and displays the previous form
    frmWomen.Hide
    frmSecondPage.Show
End Sub

Private Sub cmdShoes_Click()

    'hides this form and displays the Women shoes form
    frmWomensShoes.Show
    frmWomen.Hide
End Sub
