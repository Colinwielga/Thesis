VERSION 5.00
Begin VB.Form frmBegin 
   BackColor       =   &H0000FF00&
   Caption         =   "Money Manager"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0000FFFF&
      Caption         =   "Help!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Menu"
      Height          =   495
      Left            =   360
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdArrow 
      BackColor       =   &H00D2691E&
      DisabledPicture =   "frmBegin.frx":0000
      Height          =   1695
      Left            =   720
      Picture         =   "frmBegin.frx":9C21
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H0000FFFF&
      Caption         =   "NO."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000013&
      Height          =   3855
      Left            =   4200
      ScaleHeight     =   3795
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H0000FFFF&
      Caption         =   "YES."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H0000FF00&
      Caption         =   "by: Natalie Bly"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblInstructions1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Do you have any big bills coming up?  Are you saving for something fun like a car or tuition?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Money Manager 2005 (ProjectNataliesMoneyPlanner)
'frmBegin (frmBegin.frm)
'by Natalie Bly
'10/29/05
'This form contains the main feature of my program--a procedure that allows the user to input
'a combination of resource values (savings and earnings) and various financial obligations
'that they expect and outputs how much is "left over" to use, if there is any.  This saves
'the user from struggling to figure out how much they need to set aside.  It also gives
'the user a number to work with, a tangible value that can help them decide if they can really
'afford to make that trip to Target to buy a cd or to eat out every weekend.  It does not give
'an absolute surplus value, only an estimate, since paycheck amounts may vary, and since
'unforseen expenses have a tendency to crop up.  But it certainly can be a good place to start!
Option Explicit
Private Sub cmdArrow_Click()
    Bill = 0            'sets variables equal to zero
    BillTotal = 0
    If Surplus = 0 Then 'this is the case where the user had chosen "no" (they had no big bills), so it displays appropriately that they are only considering their paycheck at the moment
        picOutput.Print " Paycheck"; Tab(40); FormatCurrency(Paycheck)  'prints "Paycheck" and the value that the user entered as their paycheck
        Do While Bill <> -1         'uses a sentinel to signal the end of data input by user
            BillTotal = BillTotal + Bill    'sums the bill input
            Bill = InputBox("Do you have any other bills this month?  Enter the amount. When you're done, enter -1. (Note:  If you get paid every two weeks, and plan to spread out your bills between the two paychecks, divide the bill amount by 2.)", "Other Bills")
                                            'user enters as many bill values as they need to until they are done
        Loop
        Surplus = Paycheck                  'sets the Surplus value equal to Paycheck value from the "No" code to provide for the No cases and to be able to use the same variable for this loop
    End If
    Do While Bill <> -1             'uses a sentinel to signal the end of data input by user
        BillTotal = BillTotal + Bill    'sums the bill input
        Bill = InputBox("Do you have any other bills this month?  Enter the amount.  When you're done, enter -1. (Note:  If you get paid every two weeks, and plan to spread out your bills between the two paychecks, divide the bill amount by 2.)", "Other Bills")
                                    'user enters as many bill values as they need to until they are done
    Loop
    picOutput.Print "-Bill Total"; Tab(38); "-"; Tab(40); FormatCurrency(BillTotal)
                                    'prints (below the paycheck/surplus line) that we're subtracting the bill total from the Surplus followed by the value
    picOutput.Print Tab(38); "-----------------"
    Surplus = Surplus - BillTotal   'subtracts the bill total from the surplus
    picOutput.Print "Available Funds"; Tab(40); FormatCurrency(Surplus)  'prints the new surplus value after the bill total has been subtracted
    Select Case Surplus             'chooses Surplus variable to consider
        Case Is < 0                 'if the Surplus value is negative
            MsgBox "By this estimate, you have insufficient funds.  Think hard--how can you make up the difference?  Come back when you've found a solution.", , "Oops!"
                                    'displays message box that tells the user they don't have enough money
            frmBegin.Hide           'brings user back to Menu screen
            frmMenu.Show
        Case Is = 0                 'if the Surplus value is zero
            MsgBox "Wow--that is so close.  You'll be okay--as long as you don't have to buy food or basically anything else.  Keep this in mind when you're confronted with additional expenses--necessary or otherwise.  You'll have to find some way to cover that.", , "Whew!"
                                    'displays message box that warns the user that they have very little wiggle room
        Case Else   'for nonzero and non-negative Surplus values
            If (Surplus > 0 And BillTotal = 0) Then 'in the event that the surplus is positive and they have no monthly bills to consider
                MsgBox "Say!  That's swell.  What are you going to do with all of this cash?  Groceries?  Fun stuff?  Consider setting aside a certain amount or percentage of each paycheck.  Or look into investing.  There's no time like the present!", , "Hooray!"
                                    'displays a message that suggests to the user ways to use the money
            Else        'if the surplus is positive and the user had monthly bills to consider
                MsgBox "Not bad!  So this is approximately what you have to work with this month.", , "Finally!"
                                    'displays a message that informs the user that they've arrived at an estimate of how much they can spend on other things
            End If
        End Select
End Sub

Private Sub cmdEnd_Click()
    cmdYes.Enabled = True   're-enables the yes and no buttons when the user clicks on this button
    cmdNo.Enabled = True
    frmBegin.Hide           'brings the user back to the Menu screen
    frmMenu.Show
End Sub

Private Sub cmdHelp_Click()
    MsgBox "Click on the Yes button if you need to set aside money for something.  The program will use your current paycheck to estimate whether or not you will have enough to cover this expense, and will tell you approximately how much you have left over for this pay period.  Click on the No button if you only have smaller, monthly bills or none at all.  Once you enter the requested information, the Arrow button will ask for monthly bills, and incorporate this into your estimated Surplus.", , "Help"
                    'gives the user some guidance if they don't quite understand what they're supposed to do
End Sub

Private Sub cmdNo_Click()
    Savings = 0         'sets variables to zero to ensure that the calculations will turn out correctly
    Paycheck = 0
    PayPeriods = 0
    BigBill = 0
    Sum = 0
    Surplus = 0
    Paycheck = InputBox("Wow!  Lucky you!  In that case, enter the amount of your paycheck.", "Paycheck")
                        'user inputs their paycheck amount and it is stored in the Paycheck variable
    MsgBox ("Click the arrow to proceed.")  'message box intructs the user about their next step
    cmdYes.Enabled = False                  'disables the yes and no buttons to prevent strange things happening to the variable values
    cmdNo.Enabled = False
    cmdArrow.Enabled = True                 'allows the user to proceed by clicking on the arrow button, previously disabled
End Sub

Private Sub cmdYes_Click()
    Savings = 0         'sets variables to zero to ensure that the calculations will turn out correctly
    Paycheck = 0
    PayPeriods = 0
    Surplus = 0
    Sum = 0
    BigBill = 0
    BigBill = InputBox("How much do you need to cover this bill? (Estimate if you need to.)", "Big Bill Estimate")
                        'user inputs expected value of a large bill
    Savings = InputBox("Do you have savings to contribute to this bill?  How much?", "Savings")
                        'user inputs amount of savings dedicated to this bill
    Sum = Savings       'adds savings value to resource sum
    Paycheck = InputBox("Okay.  So now it's time to enter your paycheck amount.", "Paycheck")
                        'user inputs the amount of their paycheck
    PayPeriods = InputBox("Just a few more things.  How many pay periods remain before your big bill is due?", "Pay Periods Remaining")
                        'user inputs the number of remaining pay periods before the big bill is due to use this paycheck as an estimate of future paychecks that can be considered part of the their resources
    If PayPeriods > 0 Then      'if the number of pay periods remaining is greater than zero
        Sum = Sum + Paycheck + Paycheck * PayPeriods    'adds paycheck value along with estimated future paychecks to the resource sum
    Else
        Sum = Sum + Paycheck    'if there are no other pay periods remaining, then just add the paycheck value to resource sum
    End If
    picOutput.Print "Savings"; Tab(40); FormatCurrency(Savings) 'prints the text "savings" followed by the value the user input as savings
    If PayPeriods = 0 Then
        picOutput.Print "Paycheck"; Tab(38); "+"; Tab(40); FormatCurrency(Paycheck) 'prints the text "paycheck" followed by the Paycheck value
    Else
        picOutput.Print "Paycheck + Future earnings estimate"; Tab(38); "+"; Tab(40); FormatCurrency(Paycheck + (Paycheck * PayPeriods))
                            'if more pay periods do remain, then it prints the text "Paycheck+Future earnings estimate" followed by the value of the current paycheck plus the future earnings estimate
    End If
    picOutput.Print Tab(38); "-----------------"
    picOutput.Print "Available Funds"; Tab(40); FormatCurrency(Sum) 'prints the sum of Savings and Paycheck and future earnings
    picOutput.Print "Big Bill"; Tab(38); "-"; Tab(40); FormatCurrency(BigBill)  'shows the subtraction of (and value of) the big bill
    picOutput.Print Tab(38); "-----------------"
    Surplus = Sum - BigBill         'subtracts the big bill from the resource sum
    picOutput.Print "Available Funds"; Tab(40); FormatCurrency(Surplus)
    If Surplus < 0 Then             'if the bill is greater than the user's resources then
        MsgBox "Oh no!  With your resources as they are, you won't be able to cover your bill!  Quick!  Check your piggy bank!  Look under the couch cushions, check  your pockets for lost change!  Maybe you should try to get more hours at work, take out a loan to cover part or all of your bill, or talk to your parents!  You could try indentured servanthood...  Good luck.", , "Eek!  You'll be in the red!"
                                    'alerts the user to the fact that the projection indicates that they won't have enough money to cover their bill
    Else
        MsgBox "Click the arrow to proceed.", , "Continue"    'displays instructions for the user's next step
        cmdArrow.Enabled = True     'allows the user to proceed by clicking on the arrow button, previously disabled
    End If
    cmdYes.Enabled = False          'disables yes and no buttons to prevent strange things happening the the variable values to ensure correctness in the next procedures
    cmdNo.Enabled = False
             
End Sub

Private Sub Form_Load()
    cmdArrow.Enabled = False        'disables the arrow button when the form loads so
                                    'the user doesn't skip the first steps.
End Sub
