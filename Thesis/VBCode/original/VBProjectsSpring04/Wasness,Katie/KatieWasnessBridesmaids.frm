VERSION 5.00
Begin VB.Form frmBridesmaid 
   BackColor       =   &H8000000D&
   Caption         =   "BRIDESMAID ATTIRE"
   ClientHeight    =   8490
   ClientLeft      =   4575
   ClientTop       =   2880
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10815
   Begin VB.CommandButton cmdBMDress3 
      Caption         =   "Please click here to select this dress ensamble. The total cost of this dress ensamble is $100.00."
      Height          =   855
      Left            =   5880
      TabIndex        =   11
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmdBMDress2 
      Caption         =   "Please click here to select this dress ensamble. The total cost of this dress ensamble is $175.00."
      Height          =   855
      Left            =   3120
      TabIndex        =   10
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmdBMDress1 
      Caption         =   "Please click here to select this dress ensamble. The total cost of this dress ensamble is $150.00."
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton cmdUhOh 
      Caption         =   "Click here if you made a selection and have decided that you do not want that ensamble."
      Enabled         =   0   'False
      Height          =   2055
      Left            =   8640
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdboughtown 
      Caption         =   "I have bought the Bridesmaid dresses somewhere else, but I would like to input the cost of the Dress into the whole Budget."
      Height          =   3015
      Left            =   8640
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.PictureBox PicBridesmaidDress3 
      Height          =   4935
      Left            =   5880
      Picture         =   "Katie Wasness Bridesmaids.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox picBridesmaidDress2 
      Height          =   4935
      Left            =   3120
      Picture         =   "Katie Wasness Bridesmaids.frx":3F1C
      ScaleHeight     =   4875
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.PictureBox picBridesmaidDress1 
      Height          =   4935
      Left            =   360
      Picture         =   "Katie Wasness Bridesmaids.frx":90C6
      ScaleHeight     =   4875
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.PictureBox picBMDressCost 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   10395
      TabIndex        =   2
      Top             =   6600
      Width           =   10455
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Go Back to Attire Form"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   9240
      TabIndex        =   0
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label labtitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "WHICH ENSAMBLE WOULD YOU LIKE FOR YOUR BRIDESMAIDS?"
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmBridesmaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmBridesmaid(Katie Wasness Bridesmaids)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is choose the bridemaid's attire and to put the price of that into the Running Total.
Dim CostofBMDress As Single

Private Sub cmdBackTo_Click()
'This button brings you back to the Attire Form
frmAttire.Show
frmBridesmaid.Hide
End Sub

Private Sub cmdBMDress1_Click()
'this button is used if the user wants to pic dress number 1
CostofBMDress = 150#
picBMDressCost.Print "Your bridesmaids will look beautiful in this dress. The total cost of your ensamble will be "; FormatCurrency(CostofBMDress); "."
TotalCostBM = CostofBMDress * NumbersofBM
picBMDressCost.Print "The cost of all of your bridesmaids dresses added together is "; FormatCurrency(TotalCostBM); "."
RunningTotal = RunningTotal + TotalCostBM
picBMDressCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdBMDress1.Enabled = False
cmdBMDress2.Enabled = False
cmdBMDress3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdBMDress2_Click()
'this button is used if the user wants to pic dress number 2
CostofBMDress = 175#
picBMDressCost.Print "Your bridesmaids will look beautiful in this dress. The total cost of your ensamble will be "; FormatCurrency(CostofBMDress); "."
TotalCostBM = CostofBMDress * NumbersofBM
picBMDressCost.Print "The cost of all of your bridesmaids dresses added together is "; FormatCurrency(TotalCostBM); "."
RunningTotal = RunningTotal + TotalCostBM
picBMDressCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdBMDress1.Enabled = False
cmdBMDress2.Enabled = False
cmdBMDress3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdBMDress3_Click()
'this button is used if the user wants to pic dress number 3
CostofBMDress = 100#
picBMDressCost.Print "Your bridesmaids will look beautiful in this dress. The total cost of your ensamble will be "; FormatCurrency(CostofBMDress); "."
TotalCostBM = CostofBMDress * NumbersofBM
picBMDressCost.Print "The cost of all of your bridesmaids dresses added together is "; FormatCurrency(TotalCostBM); "."
RunningTotal = RunningTotal + TotalCostBM
picBMDressCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdBMDress1.Enabled = False
cmdBMDress2.Enabled = False
cmdBMDress3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdboughtown_Click()
'this button is in case an individual has boughten the bridal dress somewhere else
CostofBMDress = InputBox("Enter the amount of one Bridesmaid dress that you bought elsewhere, using a decimal point to show dollars and cents.", "Amount of Dress")
picBMDressCost.Print "I am sure the dress that you have picked out will look beautiful on you.  The total cost of your ensamble will be "; FormatCurrency(CostofBMDress); "."
TotalCostBM = CostofBMDress * NumbersofBM
picBMDressCost.Print "The cost of all of your bridesmaids dresses added together is "; FormatCurrency(TotalCostBM); "."
RunningTotal = RunningTotal + TotalCostBM
picBMDressCost.Print "Your total running total is now "; FormatCurrency(RunningTotal); "."
cmdBMDress1.Enabled = False
cmdBMDress2.Enabled = False
cmdBMDress3.Enabled = False
cmdboughtown.Enabled = False
cmdUhOh.Enabled = True

End Sub

Private Sub cmdQuit_Click()
'this button ends the program
End
End Sub



Private Sub cmdUhOh_Click()
'this button is used if the user has made a selection and wants to change it. it clears the picture box and minuses the cost of the selection from the total cost.
picBMDressCost.Cls
RunningTotal = RunningTotal - TotalCostBM
cmdBMDress1.Enabled = True
cmdBMDress2.Enabled = True
cmdBMDress3.Enabled = True
cmdboughtown.Enabled = True
End Sub

Private Sub Form_Load()
NumbersofBM = NumberofBM + 1 'to account for Maid/Matron of Honor
End Sub
