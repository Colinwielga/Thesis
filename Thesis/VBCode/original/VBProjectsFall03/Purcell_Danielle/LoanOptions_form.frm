VERSION 5.00
Begin VB.Form frmLoanOptions 
   BackColor       =   &H008080FF&
   Caption         =   "Form7"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15045
   LinkTopic       =   "Form7"
   ScaleHeight     =   11010
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreviousPage6 
      Caption         =   "Previous Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalculateLoan3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Calculate Loan Option 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   1935
   End
   Begin VB.PictureBox picLoanOption3 
      Height          =   7335
      Left            =   9960
      ScaleHeight     =   7275
      ScaleWidth      =   4755
      TabIndex        =   4
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdCalculateLoan2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Calculate Loan Option 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   1815
   End
   Begin VB.PictureBox picLoanOption2 
      Height          =   7335
      Left            =   5040
      ScaleHeight     =   7275
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdCalculateLoan1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Calculate Loan Option 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1815
   End
   Begin VB.PictureBox picLoanOption1 
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7275
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "frmLoanOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public PATH As String

Private Sub cmdCalculateLoan2_Click()
'This button will show the user the second of three loan options.
'This includes length of loan and interest rates.
'The cost of the loan will also be shown on the form.

'Dimension varibles that will be used in this form.

picLoanOption2.Cls
Dim InterestRate As Single

InterestRate = 0.045

'Figure out the risk factor of the loan by determining the
'stability of the applicant through analyzing the amount of time
'they have worked at one job.

'Dimension variables that relate to the employment factor.

Dim EmploymentRiskRate As Single, EmploymentCurrentRate As Single
Dim EmploymentPrevRate As Single, CurrentEmployment As Single, PrevEmployment As Single

'First find the rates that apply to the status of the applicant at their current job.

CurrentEmployment = frmApplicantInfo.txtYrsCurrentEmp.Text
If CurrentEmployment < 0.5 Then
    EmploymentCurrentRate = 0.02
ElseIf CurrentEmployment < 3 Then
    EmploymentCurrentRate = 0.017
ElseIf CurrentEmployment < 6 Then
    EmploymentCurrentRate = 0.014
ElseIf CurrentEmployment < 10 Then
    EmploymentCurrentRate = 0.012
ElseIf CurrentEmployment < 15 Then
    EmploymentCurrentRate = 0.007
Else
    EmploymentCurrentRate = 0.0025
End If
    
'Next, find the rates that apply to the status of the applicant at their previous job.

PrevEmployment = frmApplicantInfo.txtAppYrsPrevEmp
If PrevEmployment < 0.5 Then
    EmploymentPrevRate = 0.02
ElseIf PrevEmployment < 3 Then
    EmploymentPrevRate = 0.017
ElseIf PrevEmployment < 6 Then
    EmploymentPrevRate = 0.014
ElseIf PrevEmployment < 10 Then
    EmploymentPrevRate = 0.012
ElseIf PrevEmployment < 15 Then
    EmploymentPrevRate = 0.007
Else
    EmploymentPrevRate = 0.0025
End If

'To get the final risk rate, average the current and previous employment records.

EmploymentRiskRate = (EmploymentCurrentRate + EmploymentPrevRate) / 2

'Add the Employment Risk Rate onto the base interest rate to get a new interest rate.

InterestRate = InterestRate + EmploymentRiskRate

'Determine the risk of loaning the applicant money based on his or her stability at homes.

'Dimension variables that have to do with how long the applicant has lived in one place.

Dim HomeRiskRate As Single, CurrentHomeRate As Single, PrevHomeRate As Single
Dim CurrentHome As Single, PrevHome As Single

'First find the rates that apply to the status of the applicant at their current home.

CurrentHome = frmApplicantInfo.txtAppYrsAtCurrentAddress.Text
If CurrentHome < 0.5 Then
    CurrentHomeRate = 0.02
ElseIf CurrentHome < 3 Then
    CurrentHomeRate = 0.017
ElseIf CurrentHome < 6 Then
    CurrentHomeRate = 0.014
ElseIf CurrentHome < 10 Then
    CurrentHomeRate = 0.012
ElseIf CurrentHome < 15 Then
    CurrentHomeRate = 0.007
Else
    CurrentHomeRate = 0.0025
End If

'Then find the rates that apply to the status of the applicant at their previous home.
    
PrevHome = frmApplicantInfo.txtAppYrsAtPrevAddress.Text
If PrevHome < 0.5 Then
    PrevHomeRate = 0.02
ElseIf PrevHome < 3 Then
    PrevHomeRate = 0.017
ElseIf PrevHome < 6 Then
    PrevHomeRate = 0.014
ElseIf PrevHome < 10 Then
    PrevHomeRate = 0.02
ElseIf PrevHome < 15 Then
    PrevHomeRate = 0.007
Else
    PrevHomeRate = 0.0025
End If

'Average the current and previous home rates to determine
'the final rate to be added to the interest rate.

HomeRiskRate = (PrevHomeRate + CurrentHomeRate) / 2

'Add the home risk rate to the interest rate.

InterestRate = InterestRate + HomeRiskRate

'Now determine the the percentage of the applicant's monthly income that goes to
'paying monthly liabilities.

'Dimension variables that have to do with the monthly income of the applicant,
'a portion of the monthly income from the co-applicant (if applicable),
'and monthly payments made by the applicant.

Dim PayAbility As Single, Dependents As Single, GrossIncome As Single
Dim AppMonthlyIncome As Single
Dim CoMonthlyIncome As Single, CoGrossIncome As Single
Dim MonthlyIncome As Single

'Determine the amount of monthly income the applicant can claim after deducting
'possible expenses for dependents.

Dependents = frmApplicantInfo.txtAppNumberDependents.Text
GrossIncome = frmApplicantInfo.txtAppGrossIncome.Text
AppMonthlyIncome = FormatNumber((GrossIncome / 12) - (250 * Dependents), 2)
CoGrossIncome = frmCoApplicantInfo.txtCoAppGrossIncome.Text
CoMonthlyIncome = FormatNumber(((CoGrossIncome / 12) / 4), 2)


'The monthly income is currently equal to the applicant's monthly income and
'the applicable amount of the cosigner's income.  If the applicant does not have a
'cosigner, but the lending facility requires it due to their financial status,
'a message box will appear letting the applicant know.

MonthlyIncome = AppMonthlyIncome


'Divide monthly liabilities by monthly income to get the percentage of income
'used each month to pay bills.

PayAbility = FormatNumber((MonthlyLiabilities / MonthlyIncome), 2)

'Determine the risk factor of a loan by assigning values to the different percentages.
'If the percentage of liabilities is too high, then a cosigner may be required or no
'loan is available.  The new interest rate will also be determined in this part.


If PayAbility >= 0.9 Then
    InterestRate = (InterestRate + 0.035)
    MsgBox ("Co-signer required or no loan is available.")
End If
If PayAbility >= 0.75 And PayAbility < 0.9 Then InterestRate = InterestRate + 0.027
If PayAbility >= 0.5 And PayAbility < 0.75 Then InterestRate = InterestRate + 0.019
If PayAbility >= 0.25 And PayAbility < 0.5 Then InterestRate = InterestRate + 0.015
If PayAbility >= 0.1 And PayAbility < 0.25 Then InterestRate = InterestRate + 0.009
If PayAbility < 0.1 Then InterestRate = InterestRate + 0.004

'The program will now determine the ratio of assets to liabilities,
'also known as the quick ratio.  This ratio helps determine if the applicant is too
'far in debt to be able to afford another loan.  This is different from the above
'calculation because it takes into account loans that don't require monthly
'payments of interest.

'Dimension variables used for the calculation of the quick ratio.

Dim QuickRatio As Single

QuickRatio = TotalAssets / TotalLiabilities

If QuickRatio >= 1 Then InterestRate = InterestRate + 0.035
If QuickRatio >= 0.75 And QuickRatio < 1 Then InterestRate = InterestRate + 0.027
If QuickRatio >= 0.5 And QuickRatio < 0.75 Then InterestRate = InterestRate + 0.023
If QuickRatio >= 0.25 And QuickRatio < 0.5 Then InterestRate = InterestRate + 0.019
If QuickRatio >= 0.1 And QuickRatio < 0.25 Then InterestRate = InterestRate + 0.015
If QuickRatio < 0.1 Then
    InterestRate = InterestRate + 0.009
    MsgBox ("A Co-Signer is required or else a loan is not available.")
End If

'Open the file with the interest rate and down payment percentages for Bank One.
'The program will search the document and match the interest rate and down payment
'information.  This will help the applicant decide if the loan option would be one that
'would work for them.

'Dimension variables used to specify loan options.  Interest rate will be available through the array.


Dim Interest(1 To 9) As String
Dim DownPayment(1 To 9) As String
Dim DPPercent As String
Dim J As Integer

Open PATH & "BankTwo.Txt" For Input As #1
For J = 1 To 9
    Input #1, Interest(J), DownPayment(J)
Next J

'First we have to search the file to find the interest rate and down payment percentage.
'Dimension variables related to the search.

Dim Found As Boolean
Dim POS As Integer
POS = 0
Found = False

Do While Not Found And POS <= 9
    POS = POS + 1
    If InterestRate >= Interest(POS) Then
        Found = True
        DPPercent = DownPayment(POS)
    End If
Loop

'It is now necessary to find out how much the user wants the loan to be for.
'The program will then figure out how many months the loan period will be for,
'based on a percentage of their monthly income after payments of bills.

'The amount will be obtained through an input box.

Dim LoanAmount As Single
LoanAmount = InputBox("How much do you want the loan to be for?")

'To determine the number of months the loan needs to be over, the program will divide
'the requested amount of the loan by the applicable monthly income.

Dim LoanMonths As Single

LoanMonths = LoanAmount / AppMonthlyIncome

'Print the loan results into the picture box.

picLoanOption2.Print "The following gives you basic information about your loan option."
picLoanOption2.Print "Please review all of the information and then select a loan option."
picLoanOption2.Print "Talk to the representative of the bank to finish your loan request."
picLoanOption2.Print
picLoanOption2.Print "The interest rate is based on stability at home and in your work"
picLoanOption2.Print "place, your assets, liabilities, monthly income, and monthly payments."
picLoanOption2.Print
picLoanOption2.Print "The interest rate on your loan is "; FormatPercent(InterestRate); " annually."
picLoanOption2.Print
picLoanOption2.Print "Based on this interest rate, a down payment of "; FormatPercent(Val(DPPercent))
picLoanOption2.Print "of your loan amount is required."
picLoanOption2.Print
picLoanOption2.Print "This loan is available over a period of"; LoanMonths; "."
picLoanOption2.Print

'Next the program will compute the cost of the loan based on the future value
'of a single sum formula.  This formula takes the interest and length into account.

Dim LoanCost As Single

LoanCost = LoanAmount * (1 + InterestRate) ^ (LoanMonths / 12)
picLoanOption2.Print
picLoanOption2.Print "The total cost of this loan, including interest, will be "; FormatCurrency(LoanCost, 2)
picLoanOption2.Print

'Monthly interest payments will be required, so to determine that cost the loan amount
'will be multiplied by 1/12 of the interest rate.

Dim MonthlyLoanPayments As Single

MonthlyLoanPayments = LoanCost / LoanMonths
picLoanOption2.Print "The required monthly loan payments will be "; FormatCurrency(MonthlyLoanPayments, 2); "."

Close #1
End Sub

Private Sub cmdCalculateLoan3_Click()
'This button will show the user the last of three loan options.
'This includes length of loan and interest rates.
'The cost of the loan will also be shown on the form.

'Dimension varibles that will be used in this form.

picLoanOption3.Cls
Dim InterestRate As Single

InterestRate = 0.035

'Figure out the risk factor of the loan by determining the
'stability of the applicant through analyzing the amount of time
'they have worked at one job.

'Dimension variables that relate to the employment factor.

Dim EmploymentRiskRate As Single, EmploymentCurrentRate As Single
Dim EmploymentPrevRate As Single, CurrentEmployment As Single, PrevEmployment As Single

'First find the rates that apply to the status of the applicant at their current job.

CurrentEmployment = frmApplicantInfo.txtYrsCurrentEmp.Text
If CurrentEmployment < 0.5 Then
    EmploymentCurrentRate = 0.04
ElseIf CurrentEmployment < 3 Then
    EmploymentCurrentRate = 0.03
ElseIf CurrentEmployment < 6 Then
    EmploymentCurrentRate = 0.02
ElseIf CurrentEmployment < 10 Then
    EmploymentCurrentRate = 0.015
ElseIf CurrentEmployment < 15 Then
    EmploymentCurrentRate = 0.01
Else
    EmploymentCurrentRate = 0.005
End If
    
'Next, find the rates that apply to the status of the applicant at their previous job.

PrevEmployment = frmApplicantInfo.txtAppYrsPrevEmp.Text
If PrevEmployment < 0.5 Then
    EmploymentPrevRate = 0.04
ElseIf PrevEmployment < 3 Then
    EmploymentPrevRate = 0.03
ElseIf PrevEmployment < 6 Then
    EmploymentPrevRate = 0.02
ElseIf PrevEmployment < 10 Then
    EmploymentPrevRate = 0.015
ElseIf PrevEmployment < 15 Then
    EmploymentPrevRate = 0.01
Else
    EmploymentPrevRate = 0.005
End If

'To get the final risk rate, average the current and previous employment records.

EmploymentRiskRate = (EmploymentCurrentRate + EmploymentPrevRate) / 2

'Add the Employment Risk Rate onto the base interest rate to get a new interest rate.

InterestRate = InterestRate + EmploymentRiskRate

'Determine the risk of loaning the applicant money based on his or her stability at homes.

'Dimension variables that have to do with how long the applicant has lived in one place.

Dim HomeRiskRate As Single, CurrentHomeRate As Single, PrevHomeRate As Single
Dim CurrentHome As Single, PrevHome As Single

'First find the rates that apply to the status of the applicant at their current home.

CurrentHome = frmApplicantInfo.txtAppYrsAtCurrentAddress.Text
If CurrentHome < 0.5 Then
    CurrentHomeRate = 0.035
ElseIf CurrentHome < 3 Then
    CurrentHomeRate = 0.025
ElseIf CurrentHome < 6 Then
    CurrentHomeRate = 0.015
ElseIf CurrentHome < 10 Then
    CurrentHomeRate = 0.01
ElseIf CurrentHome < 15 Then
    CurrentHomeRate = 0.005
Else
    CurrentHomeRate = 0.0025
End If

'Then find the rates that apply to the status of the applicant at their previous home.
    
PrevHome = frmApplicantInfo.txtAppYrsAtPrevAddress.Text
If PrevHome < 0.5 Then
    PrevHomeRate = 0.035
ElseIf PrevHome < 3 Then
    PrevHomeRate = 0.025
ElseIf PrevHome < 6 Then
    PrevHomeRate = 0.015
ElseIf PrevHome < 10 Then
    PrevHomeRate = 0.01
ElseIf PrevHome < 15 Then
    PrevHomeRate = 0.005
Else
    PrevHomeRate = 0.0025
End If

'Average the current and previous home rates to determine
'the final rate to be added to the interest rate.

HomeRiskRate = (PrevHomeRate + CurrentHomeRate) / 2

'Add the home risk rate to the interest rate.

InterestRate = InterestRate + HomeRiskRate

'Now determine the the percentage of the applicant's monthly income that goes to
'paying monthly liabilities.

'Dimension variables that have to do with the monthly income of the applicant,
'a portion of the monthly income from the co-applicant (if applicable),
'and monthly payments made by the applicant.

Dim PayAbility As Single, Dependents As Single, GrossIncome As Single
Dim AppMonthlyIncome As Single
Dim CoMonthlyIncome As Single, CoGrossIncome As Single
Dim MonthlyIncome As Single

'Determine the amount of monthly income the applicant can claim after deducting
'possible expenses for dependents.

Dependents = frmApplicantInfo.txtAppNumberDependents.Text
GrossIncome = frmApplicantInfo.txtAppGrossIncome.Text
AppMonthlyIncome = FormatNumber((GrossIncome / 12) - (250 * Dependents), 2)
CoGrossIncome = frmCoApplicantInfo.txtCoAppGrossIncome.Text
CoMonthlyIncome = FormatNumber(((CoGrossIncome / 12) / 4), 2)


'The monthly income is currently equal to the applicant's monthly income and
'the applicable amount of the cosigner's income.  If the applicant does not have a
'cosigner, but the lending facility requires it due to their financial status,
'a message box will appear letting the applicant know.

MonthlyIncome = AppMonthlyIncome

'Divide monthly liabilities by monthly income to get the percentage of income
'used each month to pay bills.

PayAbility = FormatNumber((MonthlyLiabilities / MonthlyIncome), 2)

'Determine the risk factor of a loan by assigning values to the different percentages.
'If the percentage of liabilities is too high, then a cosigner may be required or no
'loan is available.  The new interest rate will also be determined in this part.


If PayAbility >= 0.9 Then
    InterestRate = (InterestRate + 0.025)
    MsgBox ("Co-signer required or no loan is available.")
End If
If PayAbility >= 0.75 And PayAbility < 0.9 Then InterestRate = InterestRate + 0.02
If PayAbility >= 0.5 And PayAbility < 0.75 Then InterestRate = InterestRate + 0.015
If PayAbility >= 0.25 And PayAbility < 0.5 Then InterestRate = InterestRate + 0.0125
If PayAbility >= 0.1 And PayAbility < 0.25 Then InterestRate = InterestRate + 0.005
If PayAbility < 0.1 Then InterestRate = InterestRate + 0.001

'The program will now determine the ratio of assets to liabilities,
'also known as the quick ratio.  This ratio helps determine if the applicant is too
'far in debt to be able to afford another loan.  This is different from the above
'calculation because it takes into account loans that don't require monthly
'payments of interest.

'Dimension variables used for the calculation of the quick ratio.

Dim QuickRatio As Single

QuickRatio = TotalAssets / TotalLiabilities

If QuickRatio >= 1 Then InterestRate = InterestRate + 0.03
If QuickRatio >= 0.75 And QuickRatio < 1 Then InterestRate = InterestRate + 0.025
If QuickRatio >= 0.5 And QuickRatio < 0.75 Then InterestRate = InterestRate + 0.025
If QuickRatio >= 0.25 And QuickRatio < 0.5 Then InterestRate = InterestRate + 0.015
If QuickRatio >= 0.1 And QuickRatio < 0.25 Then InterestRate = InterestRate + 0.0125
If QuickRatio < 0.1 Then
    InterestRate = InterestRate + 0.001
    MsgBox ("A Co-Signer is required or else a loan is not available.")
End If

picLoanOption3.Print "The following gives you basic information about your loan option."
picLoanOption3.Print "Please review all of the information and then select a loan option."
picLoanOption3.Print "Talk to the representative of the bank to finish your loan request"
picLoanOption3.Print
picLoanOption3.Print "The interest rate is based on stability at home and in your work"
picLoanOption3.Print "place, your assets, liabilities, monthly income, and monthly payments."
picLoanOption3.Print
picLoanOption3.Print "The interest rate on your loan is "; FormatPercent(InterestRate); " annually."
picLoanOption3.Print

'Open the file with the interest rate for Bank Three.
'The program will search the document and match the interest rate information.
'This will help the applicant decide if the loan option would be one that
'would work for them.

'Dimension variables used to specify loan options.  Interest rate and down
'payment will be available through the array.

Dim Interest(1 To 9) As String
Dim DPPercent As String
Dim DownPayment(1 To 9) As String
Dim J As Integer

Open PATH & "BankThree.Txt" For Input As #1
For J = 1 To 9
    Input #1, Interest(J), DownPayment(J)
Next J

'First we have to search the file to find the interest rate.

'Dimension variables related to the search.

Dim Found As Boolean
Dim POS As Integer
POS = 0
Found = False

Do While Not Found And POS <= 9
    POS = POS + 1
    If InterestRate >= Interest(POS) Then
        Found = True
        DPPercent = DownPayment(POS)
    End If
Loop
picLoanOption3.Print "Based on this interest rate, a down payment of "; FormatPercent(Val(DPPercent))
picLoanOption3.Print "of your loan amount is required."
picLoanOption3.Print



'It is now necessary to find out how much the user wants the loan to be for.
'The program will then figure out how many months the loan period will be for,
'based on a percentage of their monthly income after payments of bills.

'The amount will be obtained through an input box.

Dim LoanAmount As Single
LoanAmount = InputBox("How much do you want the loan to be for?")

'To determine the number of months the loan needs to be over, the program will divide
'the requested amount of the loan by the applicable monthly income.

Dim LoanMonths As Single

LoanMonths = LoanAmount / AppMonthlyIncome

'Print the loan results into the picture box.




picLoanOption3.Print "This loan is available over a period of"; LoanMonths; "."
picLoanOption3.Print

'Next the program will compute the cost of the loan based on the future value
'of a single sum formula.  This formula takes the interest and length into account.

Dim LoanCost As Single
Dim MonthlyLoanPayments As Single

LoanCost = LoanAmount * (1 + InterestRate) ^ (LoanMonths / 12)
picLoanOption3.Print
picLoanOption3.Print "The total cost of this loan, including interest, will be "; FormatCurrency(LoanCost, 2)
picLoanOption3.Print

'Monthly interest payments will be required, so to determine that cost the loan amount
'will be multiplied by 1/12 of the interest rate.



MonthlyLoanPayments = LoanCost / LoanMonths
picLoanOption3.Print "The required monthly loan payments will be "; FormatCurrency(MonthlyLoanPayments, 2); "."

Close #1
End Sub

Private Sub cmdPreviousPage6_Click()
'This button will take the user to the next page.
frmCompletedApp.Show
frmLoanOptions.Hide
End Sub

Private Sub cmdCalculateLoan1_Click()
'This button will show the user the first of three loan options.
'This includes length of loan and interest rates.
'The cost of the loan will also be shown on the form.

'Dimension varibles that will be used in this form.

picLoanOption1.Cls
Dim InterestRate As Single

InterestRate = 0.04

'Figure out the risk factor of the loan by determining the
'stability of the applicant through analyzing the amount of time
'they have worked at one job.

'Dimension variables that relate to the employment factor.

Dim EmploymentRiskRate As Single, EmploymentCurrentRate As Single
Dim EmploymentPrevRate As Single, CurrentEmployment As Single, PrevEmployment As Single

'First find the rates that apply to the status of the applicant at their current job.

CurrentEmployment = frmApplicantInfo.txtYrsCurrentEmp.Text
If CurrentEmployment < 0.5 Then
    EmploymentCurrentRate = 0.025
ElseIf CurrentEmployment < 3 Then
    EmploymentCurrentRate = 0.02
ElseIf CurrentEmployment < 6 Then
    EmploymentCurrentRate = 0.015
ElseIf CurrentEmployment < 10 Then
    EmploymentCurrentRate = 0.01
ElseIf CurrentEmployment < 15 Then
    EmploymentCurrentRate = 0.005
Else
    EmploymentCurrentRate = 0.0025
End If
    
'Next, find the rates that apply to the status of the applicant at their previous job.

PrevEmployment = frmApplicantInfo.txtAppYrsPrevEmp
If PrevEmployment < 0.5 Then
    EmploymentPrevRate = 0.025
ElseIf PrevEmployment < 3 Then
    EmploymentPrevRate = 0.02
ElseIf PrevEmployment < 6 Then
    EmploymentPrevRate = 0.015
ElseIf PrevEmployment < 10 Then
    EmploymentPrevRate = 0.01
ElseIf PrevEmployment < 15 Then
    EmploymentPrevRate = 0.005
Else
    EmploymentPrevRate = 0.0025
End If

'To get the final risk rate, average the current and previous employment records.

EmploymentRiskRate = (EmploymentCurrentRate + EmploymentPrevRate) / 2

'Add the Employment Risk Rate onto the base interest rate to get a new interest rate.

InterestRate = InterestRate + EmploymentRiskRate

'Determine the risk of loaning the applicant money based on his or her stability at homes.

'Dimension variables that have to do with how long the applicant has lived in one place.

Dim HomeRiskRate As Single, CurrentHomeRate As Single, PrevHomeRate As Single
Dim CurrentHome As Single, PrevHome As Single

'First find the rates that apply to the status of the applicant at their current home.

CurrentHome = frmApplicantInfo.txtAppYrsAtCurrentAddress.Text
If CurrentHome < 0.5 Then
    CurrentHomeRate = 0.025
ElseIf CurrentHome < 3 Then
    CurrentHomeRate = 0.02
ElseIf CurrentHome < 6 Then
    CurrentHomeRate = 0.015
ElseIf CurrentHome < 10 Then
    CurrentHomeRate = 0.01
ElseIf CurrentHome < 15 Then
    CurrentHomeRate = 0.005
Else
    CurrentHomeRate = 0.0025
End If

'Then find the rates that apply to the status of the applicant at their previous home.
    
PrevHome = frmApplicantInfo.txtAppYrsAtPrevAddress.Text
If PrevHome < 0.5 Then
    PrevHomeRate = 0.025
ElseIf PrevHome < 3 Then
    PrevHomeRate = 0.02
ElseIf PrevHome < 6 Then
    PrevHomeRate = 0.015
ElseIf PrevHome < 10 Then
    PrevHomeRate = 0.01
ElseIf PrevHome < 15 Then
    PrevHomeRate = 0.005
Else
    PrevHomeRate = 0.0025
End If

'Average the current and previous home rates to determine
'the final rate to be added to the interest rate.

HomeRiskRate = (PrevHomeRate + CurrentHomeRate) / 2

'Add the home risk rate to the interest rate.

InterestRate = InterestRate + HomeRiskRate

'Now determine the the percentage of the applicant's monthly income that goes to
'paying monthly liabilities.

'Dimension variables that have to do with the monthly income of the applicant,
'a portion of the monthly income from the co-applicant (if applicable),
'and monthly payments made by the applicant.

Dim PayAbility As Single, Dependents As Single, GrossIncome As Single
Dim AppMonthlyIncome As Single
Dim CoMonthlyIncome As Single, CoGrossIncome As Single
Dim MonthlyIncome As Single

'Determine the amount of monthly income the applicant can claim after deducting
'possible expenses for dependents.

Dependents = frmApplicantInfo.txtAppNumberDependents.Text
GrossIncome = frmApplicantInfo.txtAppGrossIncome.Text
AppMonthlyIncome = FormatNumber((GrossIncome / 12) - (250 * Dependents), 2)
CoGrossIncome = frmCoApplicantInfo.txtCoAppGrossIncome.Text
CoMonthlyIncome = FormatNumber(((CoGrossIncome / 12) / 4), 2)


'The monthly income is currently equal to the applicant's monthly income and
'the applicable amount of the cosigner's income.  If the applicant does not have a
'cosigner, but the lending facility requires it due to their financial status,
'a message box will appear letting the applicant know.

MonthlyIncome = AppMonthlyIncome

'Divide monthly liabilities by monthly income to get the percentage of income
'used each month to pay bills.

PayAbility = FormatNumber((MonthlyLiabilities / MonthlyIncome), 2)

'Determine the risk factor of a loan by assigning values to the different percentages.
'If the percentage of liabilities is too high, then a cosigner may be required or no
'loan is available.  The new interest rate will also be determined in this part.


If PayAbility >= 0.9 Then
    InterestRate = (InterestRate + 0.03)
    MsgBox ("Co-signer required or no loan is available.")
End If
If PayAbility >= 0.75 And PayAbility < 0.9 Then InterestRate = InterestRate + 0.025
If PayAbility >= 0.5 And PayAbility < 0.75 Then InterestRate = InterestRate + 0.0175
If PayAbility >= 0.25 And PayAbility < 0.5 Then InterestRate = InterestRate + 0.0125
If PayAbility >= 0.1 And PayAbility < 0.25 Then InterestRate = InterestRate + 0.005
If PayAbility < 0.1 Then InterestRate = InterestRate + 0.001

'The program will now determine the ratio of assets to liabilities,
'also known as the quick ratio.  This ratio helps determine if the applicant is too
'far in debt to be able to afford another loan.  This is different from the above
'calculation because it takes into account loans that don't require monthly
'payments of interest.

'Dimension variables used for the calculation of the quick ratio.

Dim QuickRatio As Single

QuickRatio = TotalAssets / TotalLiabilities

If QuickRatio >= 1 Then InterestRate = InterestRate + 0.004
If QuickRatio >= 0.75 And QuickRatio < 1 Then InterestRate = InterestRate + 0.03
If QuickRatio >= 0.5 And QuickRatio < 0.75 Then InterestRate = InterestRate + 0.025
If QuickRatio >= 0.25 And QuickRatio < 0.5 Then InterestRate = InterestRate + 0.01
If QuickRatio >= 0.1 And QuickRatio < 0.25 Then InterestRate = InterestRate + 0.005
If QuickRatio < 0.1 Then
    InterestRate = InterestRate + 0.0025
    MsgBox ("A Co-Signer is required or else a loan is not available.")
End If

'Open the file with the interest rate for Bank One.
'The program will search the document and match the interest rate
'information.  This will help the applicant decide if the loan option would be one that
'would work for them.

'Dimension variables used to specify loan options.  Interest rate and down
'payment will be available through the array.

Dim Interest(1 To 9) As String
Dim DownPayment(1 To 9) As String
Dim DPPercent As String
Dim J As Integer

Open PATH & "BankOne.Txt" For Input As #1
For J = 1 To 9
    Input #1, Interest(J), DownPayment(J)
Next J

'First we have to search the file to find the interest rate and corresponding
'down payment percentage.

'Dimension variables related to the search.

Dim Found As Boolean
Dim POS As Integer
POS = 0
Found = False

Do While Not Found And POS <= 9
    POS = POS + 1
    If InterestRate >= Interest(POS) Then
        Found = True
        DPPercent = DownPayment(POS)
    End If
Loop



'It is now necessary to find out how much the user wants the loan to be for.
'The program will then figure out how many months the loan period will be for,
'based on a percentage of their monthly income after payments of bills.

'The amount will be obtained through an input box.

Dim LoanAmount As Single
LoanAmount = InputBox("How much do you want the loan to be for?")

'To determine the number of months the loan needs to be over, the program will divide
'the requested amount of the loan by the applicable monthly income.

Dim LoanMonths As Single

LoanMonths = LoanAmount / AppMonthlyIncome

'Print the loan results into the picture box.

picLoanOption1.Print "The following gives you basic information about your loan option."
picLoanOption1.Print "Please review all of the information and then select a loan option."
picLoanOption1.Print "Talk to the representative of the bank to finish your loan request"
picLoanOption1.Print
picLoanOption1.Print "The interest rate is based on stability at home and in your work"
picLoanOption1.Print "place, your assets, liabilities, monthly income, and monthly payments."
picLoanOption1.Print
picLoanOption1.Print "The interest rate on your loan is "; FormatPercent(InterestRate); " annually."
picLoanOption1.Print
picLoanOption1.Print "Based on this interest rate, a down payment of "; FormatPercent(Val(DPPercent))
picLoanOption1.Print "of your loan amount is required."
picLoanOption1.Print
picLoanOption1.Print "This loan is available over a period of "; LoanMonths; " months."
picLoanOption1.Print

'Next the program will compute the cost of the loan based on the future value
'of a single sum formula.  This formula takes the interest and length into account.

Dim LoanCost As Single

LoanCost = LoanAmount * ((1 + InterestRate) ^ (LoanMonths / 12))
picLoanOption1.Print
picLoanOption1.Print "The total cost of this loan, including interest, will be "; FormatCurrency(LoanCost, 2)
picLoanOption1.Print

'Monthly interest payments will be required, so to determine that cost the loan amount
'will be multiplied by 1/12 of the interest rate.

Dim MonthlyLoanPayments As Single

MonthlyLoanPayments = LoanCost / LoanMonths
picLoanOption1.Print "The required monthly loan payments will be "; FormatCurrency(MonthlyLoanPayments, 2); "."

Close #1
End Sub



Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
PATH = "N:\CS130\handin\Purcell_Danielle\"
End Sub
