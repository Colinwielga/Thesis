VERSION 5.00
Begin VB.Form IncomeStatfrm 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form2"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13965
   LinkTopic       =   "Form2"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNextForm 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Click to go to next form "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10800
      Width           =   2415
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H00FF8080&
      Caption         =   "Derive Income statement"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9720
      Width           =   2415
   End
   Begin VB.PictureBox picResults3 
      Height          =   10815
      Left            =   9600
      ScaleHeight     =   10755
      ScaleWidth      =   6075
      TabIndex        =   10
      Top             =   1920
      Width           =   6135
   End
   Begin VB.TextBox txtSellingExpense 
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox txtLoan 
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox txtCostOfGoodsSold 
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox txtRevenues 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label lblOtherAdmnCosts 
      Caption         =   "Other Selling or Admnistrative Costs"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label lblInterest 
      Caption         =   "Interest on loan if any"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label lblCostOfGoods 
      Caption         =   "Cost of goods sold or service provided"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblRevenues 
      Caption         =   "     Revenues"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"IncomeStament.frx":0000
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label lblIncomeStatement 
      BackColor       =   &H00FFC0FF&
      Caption         =   " Income statement"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "IncomeStatfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Accounting basics and Income statement
'Form 3:Income statement
'Author:Patrick Niyibizi and Frankie Chan
'Date Written:October 6th 2009
'Objective:To supplement the basics of profit and loss and introduce the income statement.
Option Explicit

Private Sub cmdCompute_Click()
    Dim Revenues As Single, Costofgoodssold As Single, Othersellingcosts As Single, Interest As Single, outcome As Single    'Declare variables
    
    Revenues = txtRevenues                          'Assign varibles to textboxes
    Costofgoodssold = txtCostOfGoodsSold
    Interest = txtLoan
    Othersellingcosts = txtSellingExpense
    outcome = Revenues - (Costofgoodssold + Interest + Othersellingcosts)     'Compute Net profit or net loss
    picResults3.Print Tab(15); "Income Statement"
    picResults3.Print Tab(15); "*************************"
    
    picResults3.Print "Revenues"; Tab(35); Revenues                             'Print Revenues, cost of goods sold, Selling costs and Interest
    picResults3.Print Tab(4); "Cost of goods Sold"; Tab(35); Costofgoodssold
    picResults3.Print Tab(4); "Selling and Admnistrative Costs"; Tab(35); Othersellingcosts
    picResults3.Print Tab(4); "Interest Expense"; Tab(35); Interest
    picResults3.Print "----------------------------------------------------------------------"
    picResults3.Print
    If outcome > 0 Then
        picResults3.Print "Net Profit"; Tab(30); outcome          'Print Profit or Loss
    Else
    picResults3.Print "Net loss"; Tab(30); outcome
    End If
    
    
    
    
End Sub

Private Sub CmdNextForm_Click()     'Go to next form
    IncomeStatfrm.Hide
    BigFourFirms.Show
End Sub
