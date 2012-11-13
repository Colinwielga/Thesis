VERSION 5.00
Begin VB.Form frmCapitalBudgeting 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   900
   ClientTop       =   795
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   12750
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0000FF00&
      Caption         =   "Definitions for Terms  Above"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdNPV 
      BackColor       =   &H0000FF00&
      Caption         =   "Net Present Value"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtResidualValue 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Text            =   "15000"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txtAnnualCashFlows 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   8
      Text            =   "21500"
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdPaybackPeriod 
      BackColor       =   &H0000FF00&
      Caption         =   "Payback Period"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox txtCostofCapital 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Text            =   ".12"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtUsefulLife 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Text            =   "20"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtInitialInvestment 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Text            =   "152750"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cdmMainMenu 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Capital Budgeting Techniques:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   4560
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Capital Budgeting"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   14
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblResidualValue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the projects residual/salvage value of the project here =>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label lblAnnualCashFlows 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the annual cash flows/ annual cost savings of the project here =>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   3120
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your company's cost of capital in decimal form here =>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   5655
   End
   Begin VB.Label lblYears 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the useful life of the project in years here =>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label lblInitialInvestment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the value of the projects Initial Investment here =>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   5295
   End
End
Attribute VB_Name = "frmCapitalBudgeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Project1(Financila_Instruments.vbp)
'Form: frmCapitalBudgeting(frmCapitalBudgeting.frm)
'Author: Sean Mase and David Horn
'Date Written: March 26, 2008
'Objective:  The purpose of this form is to allow the user to conduct two diffferent capital
            'budgeting techniquesm, Payback Period and Net Present Value.  The purpose of these
            'calculations is to determine if potential project, such as a building or a piece of
            'equipment, is worth investing in.  Once, one od the calculations mentioned above is
            'performed, this form will display a message telling the user whether they should
            'invest in the proposed project.
            

Option Explicit
'declares variables, which will be used in more than one subroutine, global.
    Dim InitialInvestment As Long
    Dim CashFlows As Long

Private Sub cdmMainMenu_Click()
    'displays main menu form
    frmMainMenu.Show
    frmCapitalBudgeting.Hide
End Sub

Private Sub cmdHelp_Click()
    'diplays CBHelp form
    frmCBHelp.Show
    frmCapitalBudgeting.Hide
End Sub

Private Sub cmdNPV_Click()
    ' This button calculates the Net Present Value
    
    'declares variables
    Dim A As Integer, PV As Double, PVSum As Double, TaxRV As Single, NPV As Double
    Dim UsefulLife As Integer, COC As Single, ResidualValue As Long

    'assigns variables values
    InitialInvestment = -(txtInitialInvestment.Text)
    UsefulLife = txtUsefulLife.Text
    COC = txtCostofCapital.Text
    CashFlows = txtAnnualCashFlows.Text
    ResidualValue = txtResidualValue.Text

    'inititiates value to variables
    A = 0
    PV = 0
    PVSum = 0
    
    'for next loop that runs until the variable 'A' reaches the of the variable 'usefullife'.
    For A = 1 To UsefulLife
        Select Case A  'case statement used to do two different calculations
            Case Is = UsefulLife 'if the variable "A" equals the value of "UsefulLife" then the calc
                                 ' below is performed
                PV = (CashFlows + ResidualValue) * (1 / (1 + COC) ^ A)
                PVSum = PVSum + PV
            Case Else 'if "A" is equal to any other value than the value of "UsefulLife" then the
                      'calc below if performed
                PV = CashFlows * (1 / (1 + COC) ^ A)
                PVSum = PVSum + PV
        End Select
    Next A
    
    'Calculates NPV
    NPV = InitialInvestment + PVSum
    
    'If statement that displays two different messages.  One message if NPV >= 0
    'and another message if NPV < 0
    If NPV >= 0 Then
        MsgBox ("This investment has an NPV of " & FormatCurrency(NPV, 2) & _
            ". Since the NPV greater than or equal to 0 the project should be accepted.")
    Else
        MsgBox ("This investment has an NPV of " & FormatCurrency(NPV, 2) & _
            ". Since the NPV less than 0, the project should NOT be accepted.")
    End If
    
End Sub

Private Sub cmdPaybackPeriod_Click()
    'This button calculates the payback period
    
    'declares variables
    Dim BenchMark As Single, PaybackYR As Single
    
    'gets value for variables from the user
    InitialInvestment = txtInitialInvestment.Text
    CashFlows = txtAnnualCashFlows.Text
    BenchMark = InputBox("Enter your company's minimum payback period in years.")
    
    'calculates the payback period
    PaybackYR = InitialInvestment / CashFlows
    
    
    'If statement that displays two different messages.  One message if the payback period is >= 0
    'and another message if the payback period is < 0
    If PaybackYR <= BenchMark Then
        MsgBox ("The payback period for this project is " & FormatNumber(PaybackYR, 2) _
            & " yrs. You should accept this investment because the payback period is " _
            & "greater than or equal to your company's benchmark payback period.")
    Else
        MsgBox ("The payback period for this project is " & FormatNumber(PaybackYR, 2) _
            & " yrs. You should NOT accept this investment because the payback period " _
            & "is less than your company's benchmark payback period.")
    End If
    
End Sub


