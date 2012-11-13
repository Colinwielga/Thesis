VERSION 5.00
Begin VB.Form LoanForm 
   BackColor       =   &H0000FFFF&
   Caption         =   "BRIAN's Bank"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTitle 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   2760
      ScaleHeight     =   1935
      ScaleWidth      =   4815
      TabIndex        =   11
      Top             =   120
      Width           =   4815
   End
   Begin VB.PictureBox picBank 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   7920
      ScaleHeight     =   2655
      ScaleWidth      =   3615
      TabIndex        =   10
      Top             =   4920
      Width           =   3615
   End
   Begin VB.PictureBox picHandShake 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   8400
      ScaleHeight     =   1935
      ScaleWidth      =   2415
      TabIndex        =   9
      Top             =   2760
      Width           =   2415
   End
   Begin VB.PictureBox picFreeLoans 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   8640
      ScaleHeight     =   1815
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdApplication 
      BackColor       =   &H0000FF00&
      Caption         =   "Apply For A Loan From Brian Today"
      Height          =   975
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompare 
      BackColor       =   &H0000FF00&
      Caption         =   "Compare Brian's Loan With A Wells Fargo Loan"
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H000000FF&
      Caption         =   "End"
      Height          =   975
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdWellsRates 
      BackColor       =   &H0000FF00&
      Caption         =   "Wells Fargo's Interest Rates"
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdMonthly 
      BackColor       =   &H0000FF00&
      Caption         =   "Monthly Payments"
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotalPayments 
      BackColor       =   &H0000FF00&
      Caption         =   "Total Payments/Interest"
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdRates 
      BackColor       =   &H0000FF00&
      Caption         =   "Brian's Interest Rates"
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   3120
      ScaleHeight     =   4155
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Designed By Brian Finley"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   7680
      Width           =   1815
   End
End
Attribute VB_Name = "LoanForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Project is named LoanVBProject(FinalVBProject.vbp)
'Form Name is LoanForm(VBProject2frm.frm)
'The Author of this project is Brian Finley
'Written 3/14/04
'This Project is designed to give the user an idea of how much a one year loan will cost them
'It shows them how much interest they would be charged as well as shows them the rates of one of BRIAN'S competitors
'The second form is an application for a loan from BRIAN


Option Explicit
Dim Path As String, Ctr As Integer, MinAmount(1 To 100) As Double
Dim Rate(1 To 100) As Double, Pass As Integer, Temp As Double, Comp As Integer
Dim TempRate As Double, LoanAmount As Double, Found As Boolean
Dim X As Integer, TotalPayments As Double, TotalInterest As Double
Dim Monthly As Double, WellsMinAmount(1 To 100) As Double, WellsRate(1 To 100) As Double
Dim Z As Integer, WellsTemp As Double, WellsTempRate As Double
Dim J As Integer, WellsInterest As Double, WellsTotalPayments As Double, Savings As Double
Dim WellsMonthly As Double, MonthlySavings As Double

'takes the user to the applicaton form
Private Sub cmdApplication_Click()
ApplicationForm.Show
LoanForm.Hide
End Sub

'compares a loan from Brian with the same loan from Wells Fargo
Private Sub cmdCompare_Click()
cmdMonthly.Enabled = False
picResults.Cls
picResults.Print "Total Payments on Brian's Loan"; Tab(40); FormatCurrency(TotalPayments)
J = 0
Found = False
Do While Not Found And J < Z      'Finds Wells Interest Rate
    J = J + 1
    If LoanAmount >= WellsMinAmount(J) Then
        Found = True
        WellsInterest = WellsRate(J) * LoanAmount
        WellsTotalPayments = WellsInterest + LoanAmount
        picResults.Print "Total Payments on Wells Fargo Loan"; Tab(40); FormatCurrency(WellsTotalPayments)
    End If
Loop
picResults.Print "*********************************************************************"
Savings = WellsTotalPayments - TotalPayments
picResults.Print "Total Savings Using Brian's Loan"; Tab(40); FormatCurrency(Savings)
picResults.Print "*********************************************************************"
picResults.Print "Monthly Payments on Brian's Loan"; Tab(40); FormatCurrency(Monthly)
WellsMonthly = WellsTotalPayments / 12
picResults.Print "Monthly Payments on Wells Fargo Loan"; Tab(40); FormatCurrency(WellsMonthly)
picResults.Print "*********************************************************************"
MonthlySavings = WellsMonthly - Monthly
picResults.Print "Monthly Savings Using Brian's Loan"; Tab(40); FormatCurrency(MonthlySavings)
End Sub

'ends program
Private Sub cmdEnd_Click()
End
End Sub

'computes the monthly payments to be made on a loan from Brian
Private Sub cmdMonthly_Click()
Monthly = TotalPayments / 12
picResults.Print "*********************************************************************"
picResults.Print "Monthly Payments"; Tab(30); FormatCurrency(Monthly)
cmdWellsRates.Enabled = True
End Sub

'computes the total payments to be made on a loan from Brian
Private Sub cmdTotalPayments_Click()
cmdCompare.Enabled = False      'disables compare button to prevent errors
cmdWellsRates.Enabled = False   'disables button to prevent errors
picResults.Cls      'Clears picBox
LoanAmount = InputBox("How Much Money Would You Like A Loan For?", "Loan Amount")
Found = False
X = 0
Do While Not Found And X < Ctr      'Finds Interest Rate
    X = X + 1
    If LoanAmount >= MinAmount(X) Then
        picResults.Print "Loan Amount"; Tab(30); FormatCurrency(LoanAmount)
        picResults.Print "Interest Rate"; Tab(30); FormatPercent(Rate(X))
        Found = True 'Ends loop
    End If
Loop
If Not Found Then       'Show an error message if the amount is too small
    MsgBox "Sorry but you cannot get a loan for such a small amount", , "Error"
End If
If Found Then           'Prints out information about the desired loan
    TotalInterest = LoanAmount * Rate(X)
    picResults.Print "Total Interest"; Tab(30); FormatCurrency(TotalInterest)
    TotalPayments = TotalInterest + LoanAmount
    picResults.Print "****************************************************************************"
    picResults.Print "Total Payments"; Tab(30); FormatCurrency(TotalPayments)
    cmdMonthly.Enabled = True
End If
End Sub

'Read file into array, bubble sort to get in order, print table
Private Sub cmdRates_Click()
picResults.Cls
Open Path & "InterestRates.txt" For Input As #1
Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, MinAmount(Ctr), Rate(Ctr)
Loop
Close #1
'Bubble Sort to get the rates in order
For Pass = 1 To Ctr - 1
    For Comp = 1 To Ctr - Pass
        If MinAmount(Comp) < MinAmount(Comp + 1) Then
            Temp = MinAmount(Comp)
            MinAmount(Comp) = MinAmount(Comp + 1)
            MinAmount(Comp + 1) = Temp
            TempRate = Rate(Comp)
            Rate(Comp) = Rate(Comp + 1)
            Rate(Comp + 1) = TempRate
        End If
    Next Comp
Next Pass
'Print out table
picResults.Print Tab(15); "Brian's Loans"
picResults.Print "____________________________________________________________"
picResults.Print "Amount"; Tab(30); "Interest Rate"
picResults.Print "*******************************************************************************"
For Pass = 1 To Ctr
    picResults.Print
    picResults.Print FormatCurrency(MinAmount(Pass)); Tab(30); FormatPercent(Rate(Pass))
Next Pass
cmdTotalPayments.Enabled = True
End Sub

'reads wells rates into an array, bubble sorts, prints
Private Sub cmdWellsRates_Click()
cmdMonthly.Enabled = False
Open Path & "WellsFargoInterestRates.txt" For Input As #1
Z = 0       'Z is the Counter for the Wells Array
picResults.Cls
picResults.Print Tab(15); "Wells Fargo Loans"
picResults.Print "____________________________________________________________"
picResults.Print "Amount"; Tab(30); "Interest Rate"
picResults.Print "*******************************************************************************"
Do While Not EOF(1)     'Read file into array
    Z = Z + 1
    Input #1, WellsMinAmount(Z), WellsRate(Z)
Loop
'Bubble Sort to get the rates in order
For Pass = 1 To Z - 1
    For Comp = 1 To Z - Pass
        If WellsMinAmount(Comp) < WellsMinAmount(Comp + 1) Then
            WellsTemp = WellsMinAmount(Comp)
            WellsMinAmount(Comp) = WellsMinAmount(Comp + 1)
            WellsMinAmount(Comp + 1) = WellsTemp
            WellsTempRate = WellsRate(Comp)
            WellsRate(Comp) = WellsRate(Comp + 1)
            WellsRate(Comp + 1) = WellsTempRate
        End If
    Next Comp
Next Pass
For Pass = 1 To Z           'prints the table
    picResults.Print
    picResults.Print FormatCurrency(WellsMinAmount(Pass)); Tab(30); FormatPercent(WellsRate(Pass))
Next Pass
Close #1
cmdCompare.Enabled = True
End Sub


'When the form loads it loads the pictures and disables the buttons that will not work properly
Private Sub Form_Load()
Path = "N:\CS130\handin\Finley, Brian\"
'Path = "M:\CS130\Finley, Brian\"
'Path = "M:\CS130\VB Project\"
picTitle = LoadPicture(Path & "Title.gif")
picBank = LoadPicture(Path & "Bank.gif")
picFreeLoans = LoadPicture(Path & "FreeLoans.gif")
picHandShake = LoadPicture(Path & "HandShake.gif")
cmdMonthly.Enabled = False
cmdTotalPayments.Enabled = False
cmdWellsRates.Enabled = False
cmdCompare.Enabled = False
End Sub

