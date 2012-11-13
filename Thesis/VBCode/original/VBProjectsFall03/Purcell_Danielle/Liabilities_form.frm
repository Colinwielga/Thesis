VERSION 5.00
Begin VB.Form frmLiabilities 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form4"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13380
   LinkTopic       =   "Form4"
   ScaleHeight     =   8130
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNextPage4 
      Caption         =   "Next Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      TabIndex        =   18
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdPreviousPage3 
      Caption         =   "Previous Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   17
      Top             =   5520
      Width           =   1455
   End
   Begin VB.PictureBox picTotalMonthlyPayments 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9840
      ScaleHeight     =   915
      ScaleWidth      =   1995
      TabIndex        =   16
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalculateTotalMonthlyPayments 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Calculate Total Monthly Payments (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   1695
   End
   Begin VB.PictureBox picTotalLiabilities 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   1995
      TabIndex        =   14
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalculateTotalLiabilities 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Calculate Total Liabilities (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtMonthlyLoan 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtLoansAmt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtMonthlyRent 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtRentAmt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtMortgageMonthly 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtMortgageAmt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblMonthlyLoanAmt 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Amount Paid for Loans Outstanding per Month (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   11
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label lblLoansAmt 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Amount of Loans Outstanding (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblMonthlyRentAmt 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Rent Paid per Month (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblRentAmt 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Amount of Rent (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblMortgageMonthlyPayments 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Mortgage Monthly Payments Amount (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lblMortgage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Amount of Mortgage (in dollars)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblLiabilities 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Liabilities"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmLiabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form will be used to calculate the total liabilities of the applicant which will be compared to total assets.
'This form will also calculate monthly liabilities to compare to monthly income to see if the applicant has enough funds to support another loan payment.

Private Sub cmdCalculateTotalLiabilities_Click()
picTotalLiabilities.Cls
'Dimension variables as singles because they are representing dollar amounts.
Dim Mortgage As Single, Rent As Single, Loans As Single
Mortgage = txtMortgageAmt.Text
Rent = txtRentAmt.Text
Loans = txtLoansAmt.Text
TotalLiabilities = Mortgage + Rent + Loans
picTotalLiabilities.Print FormatCurrency(TotalLiabilities, 2)

End Sub

Private Sub cmdCalculateTotalMonthlyPayments_Click()
picTotalMonthlyPayments.Cls
'This operation will be used to calculate total monthly payments to determine whether or not is it feasible to have additional payments.
'Dimension variables as singles because they represent dollar amounts.
Dim MonthlyMortgage As Single, MonthlyRent As Single, MonthlyLoans As Single
MonthlyMortgage = txtMortgageMonthly.Text
MonthlyRent = txtMonthlyRent.Text
MonthlyLoans = txtMonthlyLoan.Text
MonthlyLiabilities = MonthlyMortgage + MonthlyRent + MonthlyLoans
picTotalMonthlyPayments.Print FormatCurrency(MonthlyLiabilities, 2)
End Sub

Private Sub cmdNextPage4_Click()
'This button will allow the user to continue onto the next page of the application,
'and also show the first page of the completed application

frmCompletedApp.picCompletedApp.Cls
frmCompletedApp.Show
frmLiabilities.Hide
frmCompletedApp.picCompletedApp.Cls
MsgBox ("Please look over the completed application and go back and make any necessary changes.")
frmCompletedApp.picCompletedApp.Print "LOAN APPLICATION"
frmCompletedApp.picCompletedApp.Print "---------------------------------------------------------------------------------------------------------"
frmCompletedApp.picCompletedApp.Print "Applicant General Information"
frmCompletedApp.picCompletedApp.Print
frmCompletedApp.picCompletedApp.Print "Name:", , , frmApplicantInfo.txtAppName.Text
frmCompletedApp.picCompletedApp.Print
frmCompletedApp.picCompletedApp.Print "Current Address:", , frmApplicantInfo.txtAppCurrentAddress.Text
frmCompletedApp.picCompletedApp.Print "Years at Current Address:", , frmApplicantInfo.txtAppYrsAtCurrentAddress.Text
frmCompletedApp.picCompletedApp.Print "City:", , , frmApplicantInfo.txtAppCurrentCity.Text
frmCompletedApp.picCompletedApp.Print "State:", , , frmApplicantInfo.txtAppCurrentState.Text
frmCompletedApp.picCompletedApp.Print "Zip Code:", , , frmApplicantInfo.txtAppCurrentZipCode.Text
frmCompletedApp.picCompletedApp.Print
frmCompletedApp.picCompletedApp.Print "Previous Address:", , frmApplicantInfo.txtAppPreviousAddress
frmCompletedApp.picCompletedApp.Print "Years at Previous Address:", frmApplicantInfo.txtAppYrsAtPrevAddress
frmCompletedApp.picCompletedApp.Print "City:", , , frmApplicantInfo.txtAppPreviousCity
frmCompletedApp.picCompletedApp.Print "State:", , , frmApplicantInfo.txtAppPreviousState
frmCompletedApp.picCompletedApp.Print "Zip Code:", , , frmApplicantInfo.txtAppPreviousZipCode
frmCompletedApp.picCompletedApp.Print
frmCompletedApp.picCompletedApp.Print "Current Employer:", , frmApplicantInfo.txtAppCurrentEmp
frmCompletedApp.picCompletedApp.Print "Years at Current Employer:", frmApplicantInfo.txtYrsCurrentEmp
frmCompletedApp.picCompletedApp.Print "Gross Income Per Year:", , FormatCurrency(Val(frmApplicantInfo.txtAppGrossIncome.Text), 2)
frmCompletedApp.picCompletedApp.Print "Number of Dependents:", , frmApplicantInfo.txtAppNumberDependents.Text
frmCompletedApp.picCompletedApp.Print "Previous Employer:", , frmApplicantInfo.txtAppPreviousEmployer.Text
frmCompletedApp.picCompletedApp.Print "Years at Previous Employer:", frmApplicantInfo.txtAppYrsPrevEmp.Text
frmCompletedApp.picCompletedApp.Print

End Sub


Private Sub cmdPreviousPage3_Click()
frmAssets.Show
frmLiabilities.Hide
End Sub

Private Sub Form_Load()
frmLiabilities.Show
frmAssets.Hide
MsgBox ("You must fill in all fields.  If a certain question does not pertain to you, please enter 'X' for a field that would require text and '0' for a field that requires numbers")
End Sub

