VERSION 5.00
Begin VB.Form frmGroomsman 
   BackColor       =   &H8000000D&
   Caption         =   "GROOMSMEN AND USHER ATTIRE"
   ClientHeight    =   8400
   ClientLeft      =   4950
   ClientTop       =   3075
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   10665
   Begin VB.PictureBox picGMUCost 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   8835
      TabIndex        =   11
      Top             =   5880
      Width           =   8895
   End
   Begin VB.CommandButton cmdGMTux3 
      Caption         =   "Please click here to select this Tux ensamble. The total cost of this Tux ensamble is $109.00."
      Height          =   735
      Left            =   6360
      TabIndex        =   10
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton cmdGMTux2 
      Caption         =   "Please click here to select this Tux ensamble. The total cost of this Tux ensamble is $99.00."
      Height          =   735
      Left            =   3240
      TabIndex        =   9
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton cmdGMTux1 
      Caption         =   "Please click here to select this Tux ensamble. The total cost of this Tux ensamble is $89.00."
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton cmdUhOh 
      Caption         =   "Click here if you made a selection and have decided that you do not want that ensamble."
      Height          =   2775
      Left            =   9240
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdboughtown 
      Caption         =   "I have bought the Tux somewhere else, but I would like to input the cost of the Tux into the Budget."
      Height          =   2175
      Left            =   9120
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.PictureBox picGMTux3 
      Height          =   2895
      Left            =   6840
      Picture         =   "Katie Wasness Groomsman.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox picGMTux2 
      Height          =   2895
      Left            =   3840
      Picture         =   "Katie Wasness Groomsman.frx":2D8F
      ScaleHeight     =   2835
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox picGMTux1 
      Height          =   2895
      Left            =   720
      Picture         =   "Katie Wasness Groomsman.frx":6073
      ScaleHeight     =   2835
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Go Back to Attire Form"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   9240
      TabIndex        =   0
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label labtitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "WHAT TUX WILL YOUR GROOMSMEN BE WEARING?"
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmGroomsman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmGroomsman(Katie Wasness Groomsman)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is choose the Groomsmen attire and put the price of that into the Running Total.
Dim CostofGMUTux As Single

Private Sub cmdBackTo_Click()
'this button is used to go back to the attire menu
frmAttire.Show
frmGroomsman.Hide
End Sub

Private Sub cmdboughtown_Click()
'this button is in case an individual has boughten the tux somewhere else
CostofGMUTux = InputBox("Enter the amount of the Tux that you bought elsewhere, using a decimal point to show dollars and cents.", "Amount of Tux")
picGMUCost.Print "Your Groomsmen and Ushers will look handsome in this Tux. The total cost of your ensamble will be "; FormatCurrency(CostofGMUTux); "."
TotalCostGMU = CostofGMUTux * NumberofGMU
picGMUCost.Print "The cost of all of your Groomsmen and Usher's Tuxes added together is "; FormatCurrency(TotalCostGMU); "."
RunningTotal = RunningTotal + TotalCostGMU
picGMUCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdGMTux1.Enabled = False
cmdGMTux2.Enabled = False
cmdGMTux3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdGMTux1_Click()
'this button is used if the user wants to pic tux number 1
CostofGMUTux = 89#
picGMUCost.Print "Your Groomsmen and Ushers will look handsome in this Tux. The total cost of your ensamble will be "; FormatCurrency(CostofGMUTux); "."
TotalCostGMU = CostofGMUTux * NumberofGMU
picGMUCost.Print "The cost of all of your Groomsmen and Usher's Tuxes added together is "; FormatCurrency(TotalCostGMU); "."
RunningTotal = RunningTotal + TotalCostGMU
picGMUCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdGMTux1.Enabled = False
cmdGMTux2.Enabled = False
cmdGMTux3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdGMTux2_Click()
'this button is used if the user wants to pic tux number 2
CostofGMUTux = 99#
picGMUCost.Print "Your Groomsmen and Ushers will look handsome in this Tux. The total cost of your ensamble will be "; FormatCurrency(CostofGMUTux); "."
TotalCostGMU = CostofGMUTux * NumberofGMU
picGMUCost.Print "The cost of all of your Groomsmen and Usher's Tuxes added together is "; FormatCurrency(TotalCostGMU); "."
RunningTotal = RunningTotal + TotalCostGMU
picGMUCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdGMTux1.Enabled = False
cmdGMTux2.Enabled = False
cmdGMTux3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdGMTux3_Click()
'this button is used if the user wants to pic tux number 3
CostofGMUTux = 109#
picGMUCost.Print "Your Groomsmen and Ushers will look handsome in this Tux. The total cost of your ensamble will be "; FormatCurrency(CostofGMUTux); "."
TotalCostGMU = CostofGMUTux * NumberofGMU
picGMUCost.Print "The cost of all of your Groomsmen and Usher's Tuxes added together is "; FormatCurrency(TotalCostGMU); "."
RunningTotal = RunningTotal + TotalCostGMU
picGMUCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdGMTux1.Enabled = False
cmdGMTux2.Enabled = False
cmdGMTux3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdQuit_Click()
'this button is to end the program
End
End Sub



Private Sub cmdUhOh_Click()
'this button is used if the user has made a selection and wants to change it. it clears the picture box and minuses the cost of the selection from the total cost.
picGMUCost.Cls
RunningTotal = RunningTotal - TotalCostGMU
cmdGMTux1.Enabled = True
cmdGMTux2.Enabled = True
cmdGMTux3.Enabled = True
cmdboughtown.Enabled = True
End Sub

Private Sub Form_Load()
NumberofGMU = NumberofGM + 1 + NumberofU 'to account for Best Man and Total
End Sub
