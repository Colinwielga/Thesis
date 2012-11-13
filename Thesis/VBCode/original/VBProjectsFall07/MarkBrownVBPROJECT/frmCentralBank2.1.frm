VERSION 5.00
Begin VB.Form frmWithdrawlsandDeposits 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Withdrawl or Deposit"
   ClientHeight    =   8190
   ClientLeft      =   465
   ClientTop       =   645
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   12540
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00008000&
      Caption         =   "Print out your receipt"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Print your receipt"
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Leave Bank"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   2055
   End
   Begin VB.PictureBox picBalance 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   3840
      ScaleHeight     =   4635
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   2640
      Width           =   4815
   End
   Begin VB.CommandButton cmdDeposit2 
      BackColor       =   &H00008000&
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Savings Deposit"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton cmdWithdrawls2 
      BackColor       =   &H00008000&
      Caption         =   "Withdrawal"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Savings Withdrawal"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdDeposit 
      BackColor       =   &H00008000&
      Caption         =   "Deposit "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Checking Deposit"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton cmdWithdrawls 
      BackColor       =   &H00008000&
      Caption         =   "Withdrawal"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Checking Withdrawal"
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   2775
   End
   Begin VB.Label lblChecking 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECKING"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblSavings 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SAVINGS"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Index           =   0
      Left            =   9480
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmWithdrawlsandDeposits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This bank system was designed and created by Mark Brown and David Bernardy

Option Explicit
Dim Balance As Double
Dim checkingdep As Double
Dim savingsdep As Double
Dim checkingwith As Double
Dim savingswith As Double
Dim J As Integer



Private Sub cmdDeposit_Click()
'Checking accound deposits

checkingdep = InputBox("How much would you like to deposit today?", "Customer Deposit Into Checking")       'Pop Up asking how much you'd like to deposit into checking
checkingbal(position + 1) = checkingbal(position + 1) + checkingdep         'Updates the account information for the array for the member

picBalance.Cls              'Clears the display

        'Displays member information in the picture box
    picBalance.Print Tab(20); "Account #"; accountnum(position + 1)
    picBalance.Print " "
    picBalance.Print firstname(position + 1); " "; lastname(position + 1)
    picBalance.Print streetadd(position + 1)
    picBalance.Print city(position + 1); ", "; state(position + 1); ", "; zipcode(position + 1)
    picBalance.Print " "
    picBalance.Print "As of "; Now; ":"
    picBalance.Print "You have "; FormatCurrency(savingsbal(position + 1)); " in your Savings Account"
    picBalance.Print "You have "; FormatCurrency(checkingbal(position + 1)); " in your Checking Account"
    picBalance.Print " "
    picBalance.Print "Thank you for banking with Central Bank."; Chr(10); "Have a nice day!"
    
End Sub

Private Sub cmdDeposit2_Click()
'savings account deposits

savingsdep = InputBox("How much would you like to deposit today?", "Customer Deposit Into Savings")         'Pop Up asking how much you'd like to deposit into your savings account
savingsbal(position + 1) = savingsbal(position + 1) + savingsdep        'Updates the account information for the array for the member

picBalance.Cls              'Clears the display

        'Displays member information in the picture box
    picBalance.Print Tab(20); "Account #"; accountnum(position + 1)
    picBalance.Print " "
    picBalance.Print firstname(position + 1); " "; lastname(position + 1)
    picBalance.Print streetadd(position + 1)
    picBalance.Print city(position + 1); ", "; state(position + 1); ", "; zipcode(position + 1)
    picBalance.Print " "
    picBalance.Print "As of "; Now; ":"
    picBalance.Print "You have "; FormatCurrency(savingsbal(position + 1)); " in your Savings Account"
    picBalance.Print "You have "; FormatCurrency(checkingbal(position + 1)); " in your Checking Account"
    picBalance.Print " "
    picBalance.Print "Thank you for banking with Central Bank."; Chr(10); "Have a nice day!"

End Sub

Private Sub cmdPrint_Click()
'Sends a customer receipt to the printer (literally, go check the printer that is connected to your computer)

Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print Tab(35); "Account #"; accountnum(ctr)
Printer.Print " "
Printer.Print Tab(20); firstname(position + 1); " "; lastname(position + 1)
Printer.Print Tab(20); streetadd(position + 1)
Printer.Print Tab(20); city(position + 1); ", "; state(position + 1); ", "; zipcode(position + 1)
Printer.Print " "
Printer.Print Tab(20); "As of "; Now; ":"
Printer.Print Tab(20); "You have "; FormatCurrency(savingsbal(position + 1)); " in your Savings Account"
Printer.Print Tab(20); "You have "; FormatCurrency(checkingbal(position + 1)); " in your Checking Account"
Printer.Print " "
Printer.Print Tab(20); "Thank you for banking with Central Bank."
Printer.Print Tab(20); "Have a great day."
Printer.EndDoc

MsgBox "Your receipt has been printed", vbExclamation, "Confirmation"       'Pop up that tells you your receipt has been printed

End Sub


Private Sub cmdWithdrawls_Click()

'Checking Accound Withdrawals

checkingwith = InputBox("How much would you like to withdraw today?", "Customer Withdrawal From Checking")       'Pop Up asking how much you'd like to withdrawal from your checking account


If checkingwith < checkingbal(position + 1) Then                'Checks that the amount you want to withdrawal from your checking account does not exceed your checking account balance
    checkingbal(position + 1) = checkingbal(position + 1) - checkingwith        'Updates the account information for the array for the member
    picBalance.Cls              'Clears the display
    
        'Displays member information in the picture box
    picBalance.Print Tab(20); "Account #"; accountnum(position + 1)
    picBalance.Print " "
    picBalance.Print firstname(position + 1); " "; lastname(position + 1)
    picBalance.Print streetadd(position + 1)
    picBalance.Print city(position + 1); ", "; state(position + 1); ", "; zipcode(position + 1)
    picBalance.Print " "
    picBalance.Print "As of "; Now; ":"
    picBalance.Print "You have "; FormatCurrency(savingsbal(position + 1)); " in your Savings Account"
    picBalance.Print "You have "; FormatCurrency(checkingbal(position + 1)); " in your Checking Account"
    picBalance.Print " "
    picBalance.Print "Thank you for banking with Central Bank."; Chr(10); "Have a nice day!"
    
    Else
        MsgBox "We're sorry, you have insufficient funds to complete this transaction.", , "Insufficient Funds"

End If

End Sub

Private Sub cmdWithdrawls2_Click()
'Saving account withdrawals

savingswith = InputBox("How much would you like to withdrawal today?", "Customer Withdrawal From Savings")       'Pop Up asking how much you'd like to withdrawal from your savings account

Select Case savingswith
Case Is < savingsbal(position + 1)                              'Checks that the amount you want to withdrawal from your savings account does not exceed your savings account balance
    savingsbal(position + 1) = savingsbal(position + 1) - savingswith           'Updates the account information for the array for the member
    
    picBalance.Cls              'Clears the display
    
        'Displays member information in the picture box
    picBalance.Print Tab(20); "Account #"; accountnum(position + 1)
    picBalance.Print " "
    picBalance.Print firstname(position + 1); " "; lastname(position + 1)
    picBalance.Print streetadd(position + 1)
    picBalance.Print city(position + 1); ", "; state(position + 1); ", "; zipcode(position + 1)
    picBalance.Print " "
    picBalance.Print "As of "; Now; ":"
    picBalance.Print "You have "; FormatCurrency(savingsbal(position + 1)); " in your Savings Account"
    picBalance.Print "You have "; FormatCurrency(checkingbal(position + 1)); " in your Checking Account"
    picBalance.Print " "
    picBalance.Print "Thank you for banking with Central Bank."; Chr(10); "Have a nice day!"
    
Case Else
    MsgBox "We're sorry, you have insufficient funds to complete this transaction.", , "Insufficient Funds"
End Select



End Sub

Private Sub Command1_Click()
'Outputs all the members of the array including the updated account

Open App.Path & "\Members.txt" For Output As #3

    For J = 1 To last
        Print #3, Chr(34); lastname(J); Chr(34); ","; Chr(34); firstname(J); Chr(34); ","; accountnum(J); ","; Chr(34); streetadd(J); Chr(34); ","; Chr(34); city(J); Chr(34); ","; Chr(34); state(J); Chr(34); ","; zipcode(J); ","; Chr(34); password(J); Chr(34); ","; savingsbal(J); ","; checkingbal(J); ","; id(J)
    Next J
    Close #3

End                                             'Exits the bank

End Sub

