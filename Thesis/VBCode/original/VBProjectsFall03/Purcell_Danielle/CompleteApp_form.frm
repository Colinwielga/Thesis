VERSION 5.00
Begin VB.Form frmCompletedApp 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form6"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13860
   LinkTopic       =   "Form6"
   ScaleHeight     =   10215
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewAssetsLiabilities 
      BackColor       =   &H00FFFF80&
      Caption         =   "View Assets and Liabilities Portion of Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdNextPage6 
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
      Left            =   12240
      TabIndex        =   3
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton cmdPreviousPage5 
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
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   8520
      Width           =   1695
   End
   Begin VB.PictureBox picCompletedApp 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   3000
      ScaleHeight     =   9315
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   360
      Width           =   9015
   End
   Begin VB.CommandButton cmdPg2CompletedApp 
      BackColor       =   &H00FFFF80&
      Caption         =   "View Completed Co-Applicant Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmCompletedApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdNextPage6_Click()
frmLoanOptions.Show
frmCompletedApp.Hide
End Sub

Private Sub cmdPg2CompletedApp_Click()
'This button will allow the applicant to view the co-applicant info to determine if the information is correct.
picCompletedApp.Cls
picCompletedApp.Print "Co-Applicant General Information"
picCompletedApp.Print
picCompletedApp.Print "Name:", , , frmCoApplicantInfo.txtCoAppName
picCompletedApp.Print
picCompletedApp.Print "Current Address:", , frmCoApplicantInfo.txtCoAppCurrentAddress.Text
picCompletedApp.Print "Years at Current Address:", , frmCoApplicantInfo.txtCoAppYrsCurrentAddress.Text
picCompletedApp.Print "City:", , , frmCoApplicantInfo.txtCoAppCurrentCity.Text
picCompletedApp.Print "State:", , , frmCoApplicantInfo.txtCoAppCurrentState.Text
picCompletedApp.Print "Zip Code:", , , frmCoApplicantInfo.txtCoAppCurrentZipCode.Text
picCompletedApp.Print
picCompletedApp.Print "Previous Address:", , frmCoApplicantInfo.txtCoAppPrevAddress.Text
picCompletedApp.Print "Years at Previous Address:", frmCoApplicantInfo.txtCoAppYrsCurrentAddress.Text
picCompletedApp.Print "City:", , , frmCoApplicantInfo.txtCoAppPrevCity.Text
picCompletedApp.Print "State:", , , frmCoApplicantInfo.txtCoAppPrevState.Text
picCompletedApp.Print "Zip Code:", , , frmCoApplicantInfo.txtCoAppPrevZipCode.Text
picCompletedApp.Print
picCompletedApp.Print "Current Employer:", , frmCoApplicantInfo.txtCoAppCurrentEmp.Text
picCompletedApp.Print "Years at Current Employer:", frmCoApplicantInfo.txtCoAppYrsCurrentEmp.Text
picCompletedApp.Print "Gross Income Per Year:", , FormatCurrency(Val(frmCoApplicantInfo.txtCoAppGrossIncome.Text), 2)
picCompletedApp.Print "Number of Dependents:", , frmCoApplicantInfo.txtCoAppNumberDependents.Text
picCompletedApp.Print "Previous Employer:", , frmCoApplicantInfo.txtCoAppPrevEmp.Text
picCompletedApp.Print "Years at Previous Employer:", frmCoApplicantInfo.txtCoAppYrsPrevEmp.Text

End Sub



Private Sub cmdPreviousPage5_Click()
frmLiabilities.Show
frmCompletedApp.Hide
End Sub

Private Sub cmdViewAssetsLiabilities_Click()
'This button will allow the user to view the assets and liabilities portion of the
'application to review.
picCompletedApp.Cls
picCompletedApp.Print "Assets"
picCompletedApp.Print
picCompletedApp.Print "Amount of Cash:", , FormatCurrency(Val(frmAssets.txtCashAmt.Text), 2)
picCompletedApp.Print "Total Value of Vehicles:", , FormatCurrency(Val(frmAssets.txtTotalVehicleValue), 2)
picCompletedApp.Print "Total Value of Real Estate:", , FormatCurrency(Val(frmAssets.txtTotalRealEstateValue), 2)
picCompletedApp.Print "Total Value of Other Assets:", FormatCurrency(Val(frmAssets.txtTotalOtherAssetsValue), 2)
picCompletedApp.Print "Total Assets:", , , FormatCurrency(Val(TotalAssets), 2)
picCompletedApp.Print
picCompletedApp.Print "Liabilities"
picCompletedApp.Print
picCompletedApp.Print "Amount of Mortgage:", , FormatCurrency(Val(frmLiabilities.txtMortgageAmt.Text), 2)
picCompletedApp.Print "Amount of Rent:", , FormatCurrency(Val(frmLiabilities.txtRentAmt.Text), 2)
picCompletedApp.Print "Amount of Loans Outstanding:", FormatCurrency(Val(frmLiabilities.txtLoansAmt.Text), 2)
picCompletedApp.Print "Total Liabilities:", , FormatCurrency(Val(TotalLiabilities), 2)
picCompletedApp.Print "Monthly Mortgage Payments:", FormatCurrency(Val(frmLiabilities.txtMortgageMonthly.Text), 2)
picCompletedApp.Print "Monthly Rent Amount:", , FormatCurrency(Val(frmLiabilities.txtMonthlyRent.Text), 2)
picCompletedApp.Print "Monthly Loan Payments:", , FormatCurrency(Val(frmLiabilities.txtMonthlyLoan.Text), 2)
picCompletedApp.Print "Total Monthly Payments:", , FormatCurrency(Val(MonthlyLiabilities), 2)
picCompletedApp.Print
End Sub


