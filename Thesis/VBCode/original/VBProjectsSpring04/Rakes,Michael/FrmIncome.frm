VERSION 5.00
Begin VB.Form FrmIncome 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdJump 
      Caption         =   "Click here to compare Phototec Inc.'s financial data to its competitors"
      Height          =   1215
      Left            =   7800
      TabIndex        =   13
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8520
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncomeV 
      Caption         =   "Click to calculate Net Income under Variable Costing"
      Height          =   1215
      Left            =   9360
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncomeA 
      Caption         =   "Click to calculate Net Income under Absorption Costing"
      Height          =   1215
      Left            =   7800
      MaskColor       =   &H0080C0FF&
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtFixed 
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtProduced 
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtSell 
      Height          =   375
      Left            =   9840
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   7455
      Left            =   360
      ScaleHeight     =   7395
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label lblFixed 
      BackColor       =   &H00FF0000&
      Caption         =   "Input the Total Fixed Costs for the quarter in dollars"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblPrice 
      BackColor       =   &H00FF0000&
      Caption         =   "Input the selling price per unit in dollars"
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   7200
      TabIndex        =   9
      Top             =   1560
      Width           =   15
   End
   Begin VB.Label lblSell 
      BackColor       =   &H00FF0000&
      Caption         =   "Input the number of units sold"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblProduced 
      BackColor       =   &H00FF0000&
      Caption         =   "Input the number of units produced"
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "FrmIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable vs. Absorption Cost Accounting(VBProject.vbp)
'FrmIncome(FrmIncome.frm)
'Mike Rakes, 3/11
'The purpose of this for is to get the number of units produced, number of units sold
'selling price, and total fixed cost for the period from the user.
'Once this information is obtained the for will print out two different income statements
'(one for each costing method) and tell whether or not the company perfomed well under each method.
'It will also explain the difference between the total net incomes under each method.
Option Explicit
'Dimensions all of the varibales
Dim UnitsP As Single, UnitsS As Single, SellPrice As Single
Dim EndInv As Single, FixedCost As Single, TotalSales As Single
Dim COGM As Single, COGS As Single, GMargin As Single, Total As String
Dim SellCosts As Single, NetIncA As Single, NetIncV As Single, CMargin As Single

Private Sub cmdIncomeA_Click()
'Sets UnitsP and UnitsS to textboxes
UnitsP = txtProduced.Text
UnitsS = txtSell.Text
'Sets Different Equations to be calculated and printed
SellPrice = txtPrice.Text
EndInv = (UnitsP - UnitsS) * FrmCosts.AC
FixedCost = txtFixed.Text
TotalSales = UnitsS * SellPrice
SellCosts = TotalSales * 0.15 + FixedCost
COGM = FrmCosts.AC * UnitsP
COGS = COGM - EndInv
GMargin = TotalSales - COGS
NetIncA = GMargin - SellCosts
'Prints a simple income statement under Absorption costing
picResults.Print "Sales"; Tab(20); FormatCurrency(TotalSales)
picResults.Print
picResults.Print "COGM"; Tab(10); FormatCurrency(COGM)
picResults.Print "End Inv."; Tab(9); "-"; Tab(10); FormatCurrency(EndInv)
picResults.Print "---------------------------------------------------------"
picResults.Print "COGS"; Tab(20); FormatCurrency(COGS)
picResults.Print
picResults.Print "Gross Margin"; Tab(20); FormatCurrency(GMargin)
picResults.Print "SG&A"; Tab(19); "-"; Tab(20); FormatCurrency(SellCosts)
picResults.Print "---------------------------------------------------------"
picResults.Print "Net Op. Inc."; Tab(20); FormatCurrency(NetIncA)
picResults.Print
picResults.Print "The Net Operating Income under absorption costing for Phototec Inc. in Quarter 1 is"; " "; FormatCurrency(NetIncA)
picResults.Print
'If income is under or over a certain amount it will
'print a statement telling you how your company performed
Select Case NetIncA
    Case Is < 0
        Total = "You have performed poorly and lost money."
    Case Is < 20000
        Total = "You have perforemed well and made money."
    Case Else
        Total = "You have performed exceptionally and made more than $20,000."
End Select
    
    picResults.Print Total
'Makes GUI more user friendly
cmdIncomeA.Enabled = False
cmdIncomeV.Enabled = True

End Sub

Private Sub cmdIncomeV_Click()
'Sets UnitsP and S to text boxes
UnitsP = txtProduced.Text
UnitsS = txtSell.Text
'Sets Different Equations to be calculated and printed
SellPrice = txtPrice.Text
TotalSales = UnitsS * SellPrice
COGM = FrmCosts.VC * UnitsP
EndInv = (UnitsP - UnitsS) * FrmCosts.VC
COGS = COGM - EndInv
FixedCost = txtFixed.Text
SellCosts = TotalSales * 0.15
CMargin = TotalSales - (COGS + SellCosts)
NetIncV = CMargin - (FixedCost + (FrmCosts.FO * UnitsP))
'prints a simple income statement under Variable Costing
picResults.Print
picResults.Print "****************************************************************************************************************"
picResults.Print "Sales"; Tab(20); FormatCurrency(TotalSales)
picResults.Print
picResults.Print "COGM"; Tab(10); FormatCurrency(COGM)
picResults.Print "End Inv."; Tab(9); "-"; Tab(10); FormatCurrency(EndInv)
picResults.Print "---------------------------------------------------------"
picResults.Print "COGS"; Tab(10); FormatCurrency(COGS)
picResults.Print
picResults.Print "SG&A"; Tab(10); FormatCurrency(SellCosts)
picResults.Print "---------------------------------------------------------"
picResults.Print "Contribution Margin"; Tab(20); FormatCurrency(CMargin)
picResults.Print "Net Op. Inc."; Tab(20); FormatCurrency(NetIncV)
picResults.Print
picResults.Print "The Net Operating Income under variable costing for Phototec Inc. in Quarter 1 is"; " "; FormatCurrency(NetIncV)
picResults.Print
'If income is under or over a certain amount it will
'print a statement telling you how your company performed
Select Case NetIncV
    Case Is < 0
        Total = "You have performed poorly and lost money."
    Case Is < 20000
        Total = "You have perforemed well and made money."
    Case Else
        Total = "You have performed exceptionally and made more than $20,000."
End Select
    picResults.Print Total
'Prints a statement about whichever income is higher
If NetIncA > NetIncV Then
    picResults.Print
    picResults.Print "****************************************************************************************************************"
    picResults.Print "Phototec Inc. has a higher Income under absorption costing than variable costing."
    picResults.Print "This is because of the difference in the ending inverntories under each costing method."
End If
If NetIncV > NetIncA Then
    picResults.Print
    picResults.Print "****************************************************************************************************************"
    picResults.Print "Phototec Inc. has a higher Incom under variable costing."
    picResults.Print "This would only happen if the fixed overhead per unit is"
    picResults.Print "negative, which is highly unlikely."
End If

'Makes GUI more user friendly
cmdJump.Enabled = True
cmdIncomeV.Enabled = False

End Sub

Private Sub cmdJump_Click()
'Jumps to another form
FrmIncome.Hide
FrmCompare.Show
End Sub

Private Sub cmdQuit_Click()
'Quits out of program
End
End Sub

Private Sub Form_Load()
'Makes GUI more user friendly
cmdJump.Enabled = False
cmdIncomeV.Enabled = False
End Sub

