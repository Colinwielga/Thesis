VERSION 5.00
Begin VB.Form frmReconciliation 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10395
   ClientLeft      =   2250
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   11145
   Begin VB.PictureBox picResults3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2760
      ScaleHeight     =   1515
      ScaleWidth      =   5475
      TabIndex        =   14
      Top             =   6960
      Width           =   5535
   End
   Begin VB.CommandButton cmdBasicInfo 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter Account Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9120
      Width           =   1935
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   5640
      ScaleHeight     =   5355
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   1440
      Width           =   5175
   End
   Begin VB.CommandButton cmdBookReconciliation 
      BackColor       =   &H0000FF00&
      Caption         =   "Reconcile Your Personal Book Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton cmdReconcile 
      BackColor       =   &H0000FF00&
      Caption         =   "Reconcile the Two Statements"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton cmdBankReconciliation 
      BackColor       =   &H0000FF00&
      Caption         =   " Reconcile Bank Statement"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   240
      ScaleHeight     =   5355
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   1440
      Width           =   5175
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0000FF00&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Personal Book Reconcilation"
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
      Left            =   7080
      TabIndex        =   13
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bank Statement Reconciliation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bank Reconciliation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   11
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblStep4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Step #4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   8760
      Width           =   855
   End
   Begin VB.Label lblSep3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Step #3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   8760
      Width           =   855
   End
   Begin VB.Label lblStep2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Step #2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label lblStep1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Step #1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   8760
      Width           =   735
   End
End
Attribute VB_Name = "frmReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Financial Tools
'From Name: frmReconciliation
'Author: David Horn & Sean Mase
'Date Written: 3-15-08
'Objective: The objective of this form is to help the user reconcile their perosnal
            'checking accounts. A Bank Reconciliation is used to make sure the user
            'has proper records of cash recipts and cash disbursements. this is
            'accomplished by comparing book records to the bank's records and
            'reconciling transations that haven't been recorded by each pary.

Option Explicit
    'declare variables glodal because they will be used by two or more subroutines.
    Dim Names As String, BankName As String, AcctNumber As String
    Dim BankCorrectBal As Single, BookCorrectBal As Single


Private Sub cmdBasicInfo_Click()
    
    'assign variables there value by getting the value from the user through inputboxes
    Names = InputBox("Enter your name.")
    BankName = InputBox("Enter the name of the bank where this account is located.")
    AcctNumber = InputBox("Enter last four digits of account number for account" _
    & "recognition.")
    
    'clears picbox
    picResults.Cls
    
    'prints header in picbox
    picResults.Print Names
    picResults.Print BankName, "Acct:"; AcctNumber
    picResults.Print "****************************************************************"
    picResults.Print
    
    'clears picbox
    picResults2.Cls
    
    'prints header in picbox
    picResults2.Print Names
    picResults2.Print BankName, "Acct:"; AcctNumber
    picResults2.Print "***************************************************************"
    picResults2.Print
    
    'dislpays instruction to go to the next step
    MsgBox ("Proceed to step #2")
End Sub

Private Sub cmdBookReconciliation_Click()
    'Declares variables
    Dim BookBal As Single
    Dim BankCredits As Single, TotalBC As Single, A As Integer
    Dim BookErrorUnder As Single, TotalBookErrorUnder As Single, B As Integer
    Dim BookAdditions As Integer, BookDeductions As Single
    Dim BankCharges As Single, TotalBankCharges As Single, C As Integer
    Dim NSF As Single, TotalNSF As Single, D As Integer
    Dim BookErrorOver As Single, TotalBookErrorOver As Single, E As Integer
    
    'assigns variable bookbal a value
    BookBal = InputBox("Enter the the balance from your peronsal books.")
    
    'prints bookbal
    picResults2.Print "Balance per pesonal books", , FormatCurrency(BookBal, 2)
    picResults2.Print Tab(5); "Add:"
    
    'Prints label for bank credits and colletctions
    picResults2.Print Tab(10); "Bank credits and collections:"
    
    'assigns value to variable
    A = 0
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    BankCredits = InputBox("Enter the value of any bank credits or collections not " _
        & "yet recorded in your books.  When all items have been entered, enter 0.")
    
    'loop to get value of input from user
    Do While BankCredits <> 0 'loop goes until user enters "0"
        A = A + 1 'initiates counter so the number of items in this category will be diplayed
        picResults2.Print , A; ")", , FormatCurrency(BankCredits, 2) 'prints the users input if it is not "0"
        TotalBC = TotalBC + BankCredits 'totals all values for this varaible
        
        'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        BankCredits = InputBox("Enter the value of any bank credits or collections " _
            & "not yet recorded in your books.  When all items have been entered, " _
            & "enter 0.")
    Loop
    
    picResults2.Print Tab(10); "Book Errors Understating Bal:"  'prints label for next set of values for this variable
    
    
    B = 0    'assigns value to variable
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    BookErrorUnder = InputBox("Enter the value of any book errors that understate " _
        & "the book balance.  When all items have been entered, enter 0.")
        
    'loop to get value of input from user
    Do While BookErrorUnder <> 0                                            'loop goes until user enters "0"
        B = B + 1                                                           'initiates counter so the number of items in this category will be diplayed
        picResults2.Print , B; ")", , FormatCurrency(BookErrorUnder, 2)     'prints the users input if it is not "0"
        TotalBookErrorUnder = TotalBookErrorUnder + BookErrorUnder          'totals all values for this varaible
        'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        BookErrorUnder = InputBox("Enter the value of any book errors that understate " _
            & "the book balance.  When all items have been entered, enter 0.")
    Loop
    
    'calculates and prints total additions
    BookAdditions = TotalBC + TotalBookErrorUnder
    picResults2.Print Tab(5); "Total Additions", , , FormatCurrency(BookAdditions, 2)
    
    'prints label
    picResults2.Print Tab(5); "Deduct:"
    
    'prints variable label
    picResults2.Print Tab(10); "Bank Charges:"
    C = 0
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    BankCharges = InputBox("Enter the value of any bank charges not yet recorded.  " _
        & "When all items have been entered, enter 0.")
        
    'loop to get value of input from user
    Do While BankCharges <> 0 'loop goes until user enters "0"
        C = C + 1                                                                   'initiates counter so the number of items in this category will be diplayed
        picResults2.Print , C; ")", , "("; FormatCurrency(BankCharges, 2); ")"      'prints the users input if it is not "0"
        TotalBankCharges = TotalBankCharges + BankCharges                           'totals all values for this varaible
        'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        BankCharges = InputBox("Enter the value of any bank charges not yet recorded." _
            & "  When all items have been entered, enter 0.")
    Loop
       
    'prints label for next set of values for this variable
    picResults2.Print Tab(10); "NSF Checks:"
    
    D = 0
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    NSF = InputBox("Enter the value of any non-sufficient funds (NSF) checks that " _
        & "have been returned to you.  When all items have been entered, enter 0.")
        
    'loop to get value of input from user
    Do While NSF <> 0
        D = D + 1                                                           'initiates counter so the number of items in this category will be diplayed
        picResults2.Print , D; ")", , "("; FormatCurrency(NSF, 2); ")"      'prints the users input if it is not "0"
        TotalNSF = TotalNSF + NSF                                           'totals all values for this varaible
        'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        NSF = InputBox("Enter the value of any non-sufficient funds (NSF) checks that " _
            & "have been returned to you.  When all items have been entered, enter 0.")
    Loop
        
    'prints label for next set of values for this variable
    picResults2.Print Tab(10); "Book Errors Overstating Bal:"
    
    'assigns value to variable
    E = 0
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    BookErrorOver = InputBox("Enter the value of any book errors that overstate your " _
        & "balance.  When all items have been entered, enter 0.")
        
    'loop to get value of input from user
    Do While BookErrorOver <> 0                                                     'loop goes until user enters "0"
        E = E + 1                                                                   'initiates counter so the number of items in this category will be diplayed
        picResults2.Print , E; ")", , "("; FormatCurrency(BookErrorOver, 2); ")"    'prints the users input if it is not "0"
        TotalBookErrorOver = TotalBookErrorOver + BookErrorOver                     'totals all values for this varaible
        
        'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        
        BookErrorOver = InputBox("Enter the value of any book errors that overstate " _
            & "your balance.  When all items have been entered, enter 0.")
    Loop
    
    'calculates and displays total deductions
    BookDeductions = TotalBankCharges + TotalBookErrorOver + TotalNSF
    picResults2.Print Tab(5); "Total Additions", , , _
         "("; FormatCurrency(BookDeductions, 2); ")"
    
    'calculates value for variable and displays correct balance afeter all entries are made
    BookCorrectBal = BookBal + BookAdditions - BookDeductions
    picResults2.Print , , , , "_________"
    picResults2.Print "Correct Cash Balance", , , FormatCurrency(BookCorrectBal, 2)
    picResults2.Print , , , , "========="
    
    
    'dislpays instruction to go to the next step
    MsgBox ("Proceed to step #4")
End Sub

Private Sub cmdMainMenu_Click()
    'dipslays MainMenu Form
    frmMainMenu.Show
    frmReconciliation.Hide
    
End Sub

Private Sub cmdBankReconciliation_Click()
    Dim BankBal As Single
    Dim DIT As Single, TotalDIT As Single, A As Integer
    Dim BErrorUnder As Single, TotalBErrorUnder As Single, B As Integer
    Dim BankAdditions As Integer, BankDeductions As Single
    Dim OutstandingChecks As Single, TotalOC As Single, C As Integer
    Dim BErrorOver As Single, TotalBErrorOver As Single, D As Integer
 
    BankBal = InputBox("Enter the balance from your bank statement.")
    picResults.Print "Balance per bank statement", , FormatCurrency(BankBal, 2)
    picResults.Print Tab(5); "Add:"
    
    'prints label for next set of values for this variable
    picResults.Print Tab(10); "Deposits in Transit:"
    
    'assigns initial value to variable
    A = 0
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    DIT = InputBox("Enter the value of any deposits in transit or undeposited " _
        & "receipts.  When all items have been entered, enter 0.")
        
    'loop to get value of input from user
    Do While DIT <> 0                                           'loop goes until user enters "0"
        A = A + 1                                               'initiates counter so the number of items in this category will be diplayed
        picResults.Print , A; ")", , FormatCurrency(DIT, 2)     'prints the users input if it is not "0"
        TotalDIT = TotalDIT + DIT                               'totals all values for this varaible
        
        'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        
        DIT = InputBox("enter the value of any deposits in transit or undeposited" _
            & "receipts.  When all items have been entered, enter 0.")
    Loop
    
    'prints label for next set of values for this variable
    picResults.Print Tab(10); "Bank Errors Understating Bal:"
    
    'assigns initial value to variable
    B = 0
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    BErrorUnder = InputBox("Enter the value of any bank errors that understate the " _
        & "bank statement balance.  When all items have been entered, enter 0.")
        
    'loop to get value of input from user
    Do While BErrorUnder <> 0                                           'loop goes until user enters "0"
        B = B + 1                                                       'initiates counter so the number of items in this category will be diplayed
        picResults.Print , B; ")", , FormatCurrency(BErrorUnder, 2)     'prints the users input if it is not "0"
        TotalBErrorUnder = TotalBErrorUnder + BErrorUnder               'totals all values for this varaible
        
        'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        BErrorUnder = InputBox("Enter the value of any bank errors that understate " _
            & "the bank statement balance.  When all items have been entered, enter 0.")
    Loop
    
    'calculates and prints total additions
    BankAdditions = TotalDIT + TotalBErrorUnder
    picResults.Print Tab(5); "Total Additions", , , FormatCurrency(BankAdditions, 2)
    
    'prints label a
    picResults.Print Tab(5); "Deduct:"
    
    
    picResults.Print Tab(10); "Outstanding Checks:"
    C = 0
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    OutstandingChecks = InputBox("Enter the value of any outstanding checks.  When " _
        & "all items have been entered, enter 0.")
        
    'loop to get value of input from user
    Do While OutstandingChecks <> 0                                                 'loop goes until user enters "0"
        C = C + 1                                                                   'initiates counter so the number of items in this category will be diplayed
        picResults.Print , C; ")", , "("; FormatCurrency(OutstandingChecks, 2); ")" 'prints the users input if it is not "0"
        TotalOC = TotalOC + OutstandingChecks                                       'totals all values for this varaible
        
       'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        OutstandingChecks = InputBox("enter the value of any outstanding checks.  " _
            & "When all items have been entered, enter 0.")
    Loop
    
    'prints label for next set of values for this variable
    picResults.Print Tab(10); "Bank Errors Overstating Bal:"
    
    'assigns initial value to variable
    D = 0
    
    'assigns anitial value to this variable, if "0" is entered the loop below will not run and the program will
    'move to get the input for the next varuable
    BErrorOver = InputBox("Enter the value of bank errors that overstate the bank " _
        & "statement balance.  When all items have been entered, enter 0.")
        
    'loop to get value of input from user
    Do While BErrorOver <> 0                                                    'loop goes until user enters "0"
        D = D + 1                                                               'initiates counter so the number of items in this category will be diplayed
        picResults.Print , D; ")", , "("; FormatCurrency(BErrorOver, 2); ")"    'prints the users input if it is not "0"
        TotalBErrorOver = TotalBErrorOver + BErrorOver                          'totals all values for this varaible
        
        'asks user to enter next value for this variable.  If the user enters 0 the loop will end
        BErrorOver = InputBox("Enter the value of any bank errors that overstate the" _
            & "bank statement balance.  When all items have been entered, enter 0.")
    Loop
    
    'calculates and displays total deductions
    BankDeductions = TotalOC + TotalBErrorOver
    picResults.Print Tab(5); "Total Additions", , , _
        "("; FormatCurrency(BankDeductions, 2); ")"
    
    
    'calculates value for variable and displays correct balance afeter all entries are made
    BankCorrectBal = BankBal + BankAdditions - BankDeductions
    picResults.Print , , , , "_________"
    picResults.Print "Correct Cash Balance", , , FormatCurrency(BankCorrectBal, 2)
    picResults.Print , , , , "========="
    
    'dislpays instruction to go to the next step
    MsgBox ("Proceed to step #3")
End Sub

Private Sub cmdReconcile_Click()
    'declatres variable
    Dim Difference As Single
    
    'Reconciles book reconciled amount and bank statement reconciled amount
    Difference = BankCorrectBal - BookCorrectBal
    
    'clears picbox
    picResults3.Cls
    
    'prints the outcome of the reconcilation calculated above
    picResults3.Print "Correct balnce per Bank Statement", , _
        FormatCurrency(BankCorrectBal, 2)
    picResults3.Print "Correct balnce per personal book", , _
        FormatCurrency(BookCorrectBal, 2)
    picResults3.Print , , , , "_________"
    picResults3.Print "Difference", , , , FormatCurrency(Difference, 2)
    picResults3.Print , , , , "========="
    
    'displays message telling the user if his or her account reconciles
    If Difference = 0 Then
        'displays message indicating the account reconciles
        MsgBox ("Your account reconciles!")
    Else
        'diplays messaged with recomended amounts to look for if account doesn't reconcile
        MsgBox ("Your account does not reconcile. (Tip: look for unrecorded items that " _
            & "have a value of " & FormatCurrency(Difference, 2) & " or recorded items" _
            & " that have a value of " & FormatCurrency((Difference / 2), 2) & " which may" _
            & " have been wrongfully subtractes or added.")
    End If
    
End Sub

Private Sub Form_Load()
    
    'displays warning that proper data type must be entered
    MsgBox ("WARNING: Make sure to enter a number when a numeric value is requested. " _
        & "If you have no value that needs to be entered make sure you enter a '0'." _
        & "If these steps are not taken the program will crash.")
End Sub
