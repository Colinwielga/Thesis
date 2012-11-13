VERSION 5.00
Begin VB.Form frmAttire 
   BackColor       =   &H8000000D&
   Caption         =   "WHAT WILL YOUR WEDDING PARTY WEAR??"
   ClientHeight    =   6255
   ClientLeft      =   6465
   ClientTop       =   4755
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   7320
   Visible         =   0   'False
   Begin VB.CommandButton cmdAttireCost 
      Caption         =   "Click to see the total cost of the Bridal Party Attire"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Go Back to Wedding Party Form"
      Height          =   615
      Left            =   3480
      TabIndex        =   6
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox picAttire 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   5175
      Left            =   2280
      ScaleHeight     =   5115
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdBridesmaids 
      Caption         =   "Choose Bridesmaids Dresses"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdGroomsman 
      Caption         =   "Choose Groomsmen and  Usher's Tuxes"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdGroom 
      Caption         =   "Choose Groom's Accessories and Tux"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdBride 
      Caption         =   "Choose Bride's Accessories and Dress"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAttire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmAttire(Katie Wasness Wedding Attire)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is to be the menu for choosing and seeing the cost of the users bridal party attire.



Private Sub cmdAttireCost_Click()
'this button is used to see the entire cost for the bridal party attire
picAttire.Cls
TotalCostofAttire = CostofDress + TotalCostBM + CostofTux + TotalCostGMU
picAttire.Print "The Total Cost of the Wedding Party Attire is "; FormatCurrency(TotalCostofAttire); "."
picAttire.Print ""
picAttire.Print "The Cost of the Bride's Attire is "; FormatCurrency(CostofDress); "."
picAttire.Print ""
picAttire.Print "The Cost of the Groom's Attire is "; FormatCurrency(CostofTux); "."
picAttire.Print ""
picAttire.Print "The Total Cost of the Bridesmaid's Attire is "; FormatCurrency(TotalCostBM); "."
picAttire.Print ""
picAttire.Print "The Total Cost of the Groomsmen and Usher's Attire is "; FormatCurrency(TotalCostGMU)
picAttire.Print ""
picAttire.Print "The cost of your Wedding to this point is "; FormatCurrency(RunningTotal); "."

End Sub

Private Sub cmdBackTo_Click()
'this button is used to go back to the wedding party menu
FrmWeddingParty.Show
frmAttire.Hide
End Sub

Private Sub cmdBride_Click()
'this button is used to to the form to choose the bridal attire
frmBride.Show
frmAttire.Hide
cmdGroom.Enabled = True

End Sub

Private Sub cmdBridesmaids_Click()
'this button is used to go to the form to choose the Bridesmaids Attire
frmBridesmaid.Show
frmAttire.Hide
cmdGroomsman.Enabled = True
End Sub

Private Sub cmdGroom_Click()
'this button is used to go to the form to choose the grooms attire
frmGroom.Show
frmAttire.Hide
cmdBridesmaids.Enabled = True
End Sub

Private Sub cmdGroomsman_Click()
'this button is used to go to the form to choose the groomsmen's attire
frmGroomsman.Show
frmAttire.Hide
cmdAttireCost.Enabled = True
End Sub

Private Sub cmdQuit_Click()
'this button is used to quit the program
End
End Sub
