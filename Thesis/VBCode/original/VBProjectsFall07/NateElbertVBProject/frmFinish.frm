VERSION 5.00
Begin VB.Form frmFinish 
   Caption         =   "Properties of Estimated Tax Return"
   ClientHeight    =   11550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19875
   LinkTopic       =   "Form1"
   ScaleHeight     =   11550
   ScaleWidth      =   19875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7440
      TabIndex        =   2
      Top             =   9600
      Width           =   4935
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Print Estimated Tax Return"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.PictureBox picResultsFinish 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   1800
      ScaleHeight     =   7875
      ScaleWidth      =   16035
      TabIndex        =   0
      Top             =   1440
      Width           =   16095
   End
End
Attribute VB_Name = "frmFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFinish_Click()
    
    picResultsFinish.Print Tab(5); "Name";                                          'Personal information of user is displayed
    picResultsFinish.Print Tab(15); "Address";
    picResultsFinish.Print Tab(25); "City";
    picResultsFinish.Print Tab(35); "State";
    picResultsFinish.Print Tab(45); "Zip Code"
    picResultsFinish.Print "***************************************************************************************************************"
    picResultsFinish.Print Tab(5); UserName; Tab(15); Address; Tab(25); City; Tab(35); State; Tab(45); ZIPcode
    picResultsFinish.Print " "
    picResultsFinish.Print "Pertinent information used to estimate tax return."                             'Information regarding users taxes is displayed
    picResultsFinish.Print "Wages for tax period: "; Tab(30); FormatCurrency(Wages(CTR))
    picResultsFinish.Print "Taxable Interest for tax period:"; Tab(30); FormatCurrency(TaxableInterest)
    picResultsFinish.Print "Unemployment Compensation for tax period:"; Tab(30); FormatCurrency(UnemploymentCompensation)
    picResultsFinish.Print "Ajusted Gross Income for tax period:"; Tab(30); FormatCurrency(AGI)
    picResultsFinish.Print "Deductions: "; Tab(30); "$7800.00"
    picResultsFinish.Print "Taxable income for tax period:"; Tab(30); FormatCurrency(TaxableIncome)
    picResultsFinish.Print "Federal income taxes withheld:"; Tab(30); FormatCurrency(IncomeTax)
    picResultsFinish.Print "Earned Income Credit:"; Tab(30); FormatCurrency(EIC)
    picResultsFinish.Print "Total Payments to Income Taxes:"; Tab(30); FormatCurrency(TotalPayments)
    picResultsFinish.Print "Total Tax:"; Tab(30); FormatCurrency(Tax(CTR))
    picResultsFinish.Print "Estimated refund amount:"; Tab(30); FormatCurrency(Refund)                                      'Estimated amount of users refund is displayed
    picResultsFinish.Print "Estimated taxes owed:"; Tab(30); FormatCurrency(Owe)                                            'Estimated amount of what user owes in taxes is displayed
    
End Sub

Private Sub cmdQuit_Click()
    End                                                                                                         'Button ends the program
End Sub

