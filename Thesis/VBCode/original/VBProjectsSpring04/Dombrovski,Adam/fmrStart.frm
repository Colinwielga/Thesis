VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H80000009&
   Caption         =   "Home Page"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShop 
      BackColor       =   &H000000FF&
      Caption         =   "Shop our Store"
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   1200
      Picture         =   "fmrStart.frx":0000
      Top             =   1680
      Width           =   2160
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : AdvoCare Store (Dombrovski,Adam.vbp)
'Form Name : frmStart (Home Page)
'Author: Adam Dombrovski
'Date Written: March 15, 2004
'Purpose: The purpose of this form is to introduce the name of the
    'company who's products will be used in the program to the user.
    'The purpose of the project is to allow a user to learn about the
    'different products that AdvoCare has and to simulate a purchase.
    'There is a module that is used to declare all the variables used
    'in the program
    

Private Sub cmdShop_Click()
frmStart.Hide
frmChoose.Show

End Sub

Private Sub Form_Load()
MsgBox "All of the products, prices and pictures are factual and have been taken from AdvoCare's website.  If you would like to find out more visit www.advocare.com or contact your local distributor...me, Adam Dombrovski"
cmdShop.Enabled = True
runningTotal = 0
ctr = 0
End Sub

