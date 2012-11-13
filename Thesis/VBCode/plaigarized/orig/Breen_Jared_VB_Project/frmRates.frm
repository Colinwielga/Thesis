VERSION 5.00
Begin VB.Form frmRates 
   Caption         =   "Rates"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRate 
      Height          =   375
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   16
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox picDemand 
      Height          =   375
      Left            =   7320
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox picSatisfaction 
      Height          =   375
      Left            =   7320
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.PictureBox picRevenue 
      Height          =   375
      Left            =   6480
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   3120
      Width           =   2175
   End
   Begin VB.PictureBox picExpenses 
      Height          =   375
      Left            =   6480
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   3840
      Width           =   2175
   End
   Begin VB.PictureBox picProfits 
      Height          =   375
      Left            =   6480
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.PictureBox picYear 
      Height          =   375
      Left            =   8040
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Operations"
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "Adjust Rates"
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblRate 
      Caption         =   "Current Rate (/MW):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblDemand 
      Caption         =   "Power Demand (MW/h)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lblSatisfaction 
      Caption         =   "Popular Satisfaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lblRevenue 
      Caption         =   "Annual Revenue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblCosts 
      Caption         =   "Annual Expenses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label lblProfit 
      Caption         =   "Annual Profits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lbl1 
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblRates 
      Caption         =   "Rates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: PowerSim2010
'Form name: frmRates
'Author: Jared Breen
'Date Written: February 18, 2010
'Objective:
'This form allows the player to adjust the rates charged by the company for electricity.  Raising rates decreases public satisfaction
'   and reduces electricity demand, but increases the amount of money brought in per MW/h sold.  Lowering rates has the opposite
'   effect, albeit the gain in satisfaction is reduced to prevent the player from raising and lowering every few years without
'   penalty.  Use of this screen early in the game is usually necessary, especially to run a clean energy strategy.

Option Explicit
Private Sub cmdAdjust_Click()
    'Set up variable and get input from the user
    Dim Temp As Single
    Temp = InputBox("Enter 1 to raise rates or -1 to lower rates.  Anything else will cancel this operation.  All rate changes are in increments of 10 cents and can only be done once a year.")
    Select Case Temp
        'Raising rates increases the rate, decreases satisfaction and decreases demand
        Case Is = 1
            Rate = Rate + 0.1
            Satisfaction = Satisfaction - 5
            Demand = Demand * 0.98
            'Adjusting non-directly affected variables
            Balance = Production - Demand
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            'Refreshing stats and disabling rate changes for the rest of the year
            frmRates.picDemand.Cls
            frmRates.picSatisfaction.Cls
            frmRates.picRevenue.Cls
            frmRates.picProfits.Cls
            frmRates.picRate.Cls
            frmRates.picDemand.Print Demand
            frmRates.picSatisfaction.Print Satisfaction
            frmRates.picRevenue.Print FormatCurrency(Revenue)
            frmRates.picProfits.Print FormatCurrency(Profit)
            frmRates.picRate.Print FormatCurrency(Rate)
            cmdAdjust.Enabled = False
        Case Is = -1
            'Raising rates decreases the rate, increases satisfaction and increases demand
            Rate = Rate - 0.1
            Satisfaction = Satisfaction + 3
            Demand = Demand * 1.02
            'Adjusting non-directly affected variables
            If Balance >= 0 Then
                Revenue = Demand * Rate
            Else
                Revenue = (Demand + Balance) * Rate
            End If
            Profit = Revenue - Expenses
            'Refreshing stats and disabling rate changes for the rest of the year
            frmRates.picDemand.Cls
            frmRates.picSatisfaction.Cls
            frmRates.picRevenue.Cls
            frmRates.picProfits.Cls
            frmRates.picRate.Cls
            frmRates.picDemand.Print Demand
            frmRates.picSatisfaction.Print Satisfaction
            frmRates.picRevenue.Print FormatCurrency(Revenue)
            frmRates.picProfits.Print FormatCurrency(Profit)
            frmRates.picRate.Print FormatCurrency(Rate)
            cmdAdjust.Enabled = False
        Case Else
            'User cancels operation
            MsgBox ("That is not a valid option.  Operation canceled.")
    End Select
End Sub

Private Sub cmdReturn_Click()
    'Switches forms and refreshes the main screen
    frmRates.Hide
    frmMain.Show
    frmMain.picYear.Cls
    frmMain.picDemand.Cls
    frmMain.picProduction.Cls
    frmMain.picBalance.Cls
    frmMain.picSatisfaction.Cls
    frmMain.picFunds.Cls
    frmMain.picRevenue.Cls
    frmMain.picExpenses.Cls
    frmMain.picProfits.Cls
    frmMain.picYear.Print Year
    frmMain.picDemand.Print Demand
    frmMain.picProduction.Print Production
    frmMain.picBalance.Print Balance
    frmMain.picSatisfaction.Print Satisfaction
    frmMain.picFunds.Print FormatCurrency(Funds)
    frmMain.picRevenue.Print FormatCurrency(Revenue)
    frmMain.picExpenses.Print FormatCurrency(Expenses)
    frmMain.picProfits.Print FormatCurrency(Profit)
End Sub
