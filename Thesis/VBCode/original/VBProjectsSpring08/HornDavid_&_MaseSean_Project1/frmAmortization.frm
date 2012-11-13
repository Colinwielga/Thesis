VERSION 5.00
Begin VB.Form frmAmortization 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   9690
   ClientLeft      =   2250
   ClientTop       =   795
   ClientWidth     =   11325
   LinkTopic       =   "Form2"
   ScaleHeight     =   9690
   ScaleWidth      =   11325
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Why do you need an amortization schedule?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox picResults9 
      Height          =   615
      Left            =   2520
      ScaleHeight     =   555
      ScaleWidth      =   6675
      TabIndex        =   17
      Top             =   4200
      Width           =   6735
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000FF00&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0000FF00&
      Caption         =   "Click to Return to main menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalculateAmort 
      BackColor       =   &H0000FF00&
      Caption         =   "Click to display Amortization table"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7800
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2520
      ScaleHeight     =   2715
      ScaleWidth      =   7035
      TabIndex        =   8
      Top             =   4800
      Width           =   7095
      Begin VB.PictureBox picResults3 
         Height          =   2295
         Left            =   240
         ScaleHeight     =   2235
         ScaleWidth      =   3675
         TabIndex        =   16
         Top             =   240
         Width           =   3735
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2775
         Left            =   4440
         TabIndex        =   15
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.TextBox txtInterest 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7920
      TabIndex        =   3
      Text            =   ".065"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtAmtBorrowed 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7920
      TabIndex        =   2
      Text            =   "150000"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtPaymentsPerYR 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7920
      TabIndex        =   1
      Text            =   "12"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtLengthYRS 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7920
      TabIndex        =   0
      Text            =   "15"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mortgage Payment and Amortization Table Calculator "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   14
      Top             =   120
      Width           =   9495
   End
   Begin VB.Label lblPaymentValue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Value of periodic payments equaks =>"
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
      Left            =   4320
      TabIndex        =   12
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Label lblLengthYRS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the length of your mortgage in years here =>"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Label lblPaymentsPerYR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the number of payments made per year here =>"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   2640
      Width           =   5055
   End
   Begin VB.Label lblInterest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your interest rate as decimal value here =>"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label lblAmtBorrowed 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the total amount borrowed here =>"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   2160
      Width           =   3975
   End
End
Attribute VB_Name = "frmAmortization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Project1(Financila_Instruments.vbp)
'Form: frmAmotization(frmAmortization.frm)
'Author: Sean Mase and David Horn
'Date Written: March 26, 2008
'Objective: The purpose of this form is to provide the user with the value of
            ' of their payments for loans.  Also, this form prints out an
            ' amortization schedule which is helpful for mortgage loans
            ' because the interest on these loans is tax deductable.

Option Explicit

Private Sub cmdAmort_Click()
    'brings the user back to the MainMenu form
    frmAmortization.Show
    frmHomeForm.Hide
End Sub


Private Sub cmdCalculateAmort_Click()
    'This button calculates and displays the payment and amortization schedule.
    
    
    'declare variables
    Dim Interest As Single, PaymentAmount As Single, Principal As Single
    Dim NumPaymentsPerYr As Integer, Years As Integer, CTR As Integer
    Dim AmtOutstanding As Single, Rate As Single
    
    'assign anitial values for variables
    CTR = 0
    Rate = txtInterest.Text
    AmtOutstanding = txtAmtBorrowed.Text
    NumPaymentsPerYr = txtPaymentsPerYR
    Years = txtLengthYRS
    
    
    PaymentAmount = (AmtOutstanding * (Rate / NumPaymentsPerYr)) / _
        (1 - ((1 + (Rate / NumPaymentsPerYr)) ^ (-NumPaymentsPerYr * Years)))
    
    'clear printboxes
    picResults2.Cls
    picResults3.Cls
    picResults9.Cls
    
    'Prints header for amortization schedule into picResults9
    picResults9.Print , "Part of payment", "Part of Payment"
    picResults9.Print "Payment #", "Representing Interest", "Representing Principal",
    picResults9.Print "Amount Oustanding"
    picResults9.Print "**************************************************************";
    picResults9.Print "**************************************************"
    
    
    'Prints the payment amount for the loan
    picResults2.Print FormatCurrency(PaymentAmount, 2)
    
    'Prints the intial amount of the loan
    picResults3.Print , , , , , FormatCurrency(AmtOutstanding, 2)
    
    'the Do Whole Loop below calculates and prints the amortization schedule.
    Do While AmtOutstanding > 0.1                               'stops the calculation when the amount outstansing is less than $0.10. if their is a remainder is due rounding errors
          Interest = (Rate / NumPaymentsPerYr) * AmtOutstanding 'calculates part of payment that represents iterest
          Principal = PaymentAmount - Interest                  'calculates amount of payment that represents principal
          AmtOutstanding = AmtOutstanding - Principal           'reduces amount outstanfing
          CTR = CTR + 1
          picResults3.Print Tab(5); CTR, FormatCurrency(Interest, 2), ,
          picResults3.Print FormatCurrency(Principal, 2), ,
          picResults3.Print FormatCurrency(AmtOutstanding, 2)
    Loop
     
End Sub

Private Sub cmdClear_Click()
    'Clears picboxes
    picResults3.Cls
    picResults2.Cls
    
End Sub

Private Sub cmdMainMenu_Click()
    'displays the MainMenu form
    frmMainMenu.Show
    frmAmortization.Hide
End Sub

Private Sub Command1_Click()
    'displays message of importance of amortization schedules
    MsgBox ("There are two reasons why it is important to calculate an amortization " _
        & "schedule. 1.) If you are amortizing a mortgage loan, it is beneficial to " _
        & "know what amount of your payments represents interest because this interest" _
        & " is tax deductable. 2.) Amortization shcedules help emphasize the fact that" _
        & " with loans a sizeable amount iterest is going to be paid to the lender.")

End Sub

Private Sub Form_Resize()
    'the code below was taken from a VB exmple found at N:\Classes\CS130\Vb_Examples.
    'the title of the example is "scrollbarTry.vbp."
    
    'To Position picResults3 in pisResults
    picResults3.Height = 35525
    picResults3.Width = picResults.Width - VScroll1.Width
    picResults3.Move 0, 0
    
    
    'to position the scrollbar
    VScroll1.Top = 0
    VScroll1.Height = picResults.Height
    VScroll1.Left = picResults.Width - VScroll1.Width
    
    'set the max properties of the scrollbar
     VScroll1.Max = picResults3.Height - picResults.Height
    
End Sub

Private Sub VScroll1_Change()
    'the code below was taken from a VB exmple found at N:\Classes\CS130\Vb_Examples.
    'the title of the example is "scrollbarTry.vbp."
    
    'allows for the scoll bar to scroll
    picResults3.Top = -VScroll1.Value
    
    'makes the function in the buttim cmdCalculateAmort repeat so the print in the
    'picbox doen't disapear when scrolled.
    Call cmdCalculateAmort_Click
End Sub
