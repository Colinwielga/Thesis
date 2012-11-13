VERSION 5.00
Begin VB.Form frmPaymentsAndTax 
   Caption         =   "Payments and Tax Section of Tax Return"
   ClientHeight    =   9660
   ClientLeft      =   2985
   ClientTop       =   2100
   ClientWidth     =   19590
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   19590
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   13320
      TabIndex        =   32
      Top             =   6000
      Width           =   4575
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Return to Previous Page"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8040
      TabIndex        =   31
      Top             =   6000
      Width           =   4935
   End
   Begin VB.CommandButton cmdContinue3 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2400
      TabIndex        =   30
      Top             =   6000
      Width           =   5175
   End
   Begin VB.PictureBox picResultsOwe 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15960
      ScaleHeight     =   675
      ScaleWidth      =   3435
      TabIndex        =   29
      Top             =   4800
      Width           =   3495
   End
   Begin VB.PictureBox picResultsRefund 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15960
      ScaleHeight     =   675
      ScaleWidth      =   3435
      TabIndex        =   24
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CommandButton cmdRefund 
      Caption         =   "Compute Refund  or Taxes Owed"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   23
      Top             =   3840
      Width           =   3975
   End
   Begin VB.PictureBox picResultsFedTax 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15960
      ScaleHeight     =   555
      ScaleWidth      =   3435
      TabIndex        =   19
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdFedTax 
      Caption         =   "Load from W2"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   18
      Top             =   240
      Width           =   3975
   End
   Begin VB.PictureBox picResultsTax 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15960
      ScaleHeight     =   675
      ScaleWidth      =   3435
      TabIndex        =   17
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton cmdComputeTax 
      Caption         =   "Compute Tax"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   16
      Top             =   2880
      Width           =   3975
   End
   Begin VB.PictureBox picResultsTotalPayments 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15960
      ScaleHeight     =   675
      ScaleWidth      =   3435
      TabIndex        =   15
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CommandButton cmdComputePayments 
      Caption         =   "Compute total payments"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   14
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtEIC 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   13
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Taxes Owed"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   33
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label Label14 
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   28
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "If line 10 is larger than line 9 then this is what you owe."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   27
      Top             =   4800
      Width           =   6135
   End
   Begin VB.Label Label12 
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   26
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "Refund"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   22
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblLine11 
      Caption         =   "If line 9 is larger than line 10  then this is your refund."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   21
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Label lbl15 
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label blb14 
      Caption         =   "Tax. This is computed from line 6 on the previous page using the IRS' progressive tax brackets."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   12
      Top             =   2880
      Width           =   6495
   End
   Begin VB.Label lbl13 
      Caption         =   "Add lines 7 and 8. These are your total payments."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Label lbl11 
      Caption         =   "Earned income credit (EIC)."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   10
      Top             =   960
      Width           =   6375
   End
   Begin VB.Label lbl10 
      Caption         =   "Federal income tax withheld from box 2 of your Form W-2."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label9 
      Caption         =   "Payments and Tax"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   7
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmPaymentsAndTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdContinue3_Click()
    frmFinish.Show                                                                  'Button hides Payments and Tax form and shows the Finish form
    frmPaymentsAndTax.Hide
End Sub



Private Sub cmdPrevious_Click()
    frmPaymentsAndTax.Hide
    frmIncome.Show                                                                  'Button returns to the previous page
    
End Sub

Private Sub cmdQuit_Click()
End                                                                                 'Button ends the program
End Sub

Private Sub cmdRefund_Click()
    If TotalPayments > IncomeTax Then                                               'Button determines amount of refund due to user or tax owed by user and displays such amount
    Refund = TotalPayments - IncomeTax
        picResultsRefund.Print Refund
        picResultsOwe.Print "None"
    Else
        Owe = IncomeTax - TotalPayments
            picResultsOwe.Print Owe
            picResultsRefund.Print "None"
    End If
    
End Sub




Private Sub cmdComputePayments_Click()
    EIC = txtEIC.Text                                                                  'Earned Income Credit is read from a textbox
    TotalPayments = Tax(CTR) + EIC                                                      'Total payments made towards taxes are computed and displayed
    picResultsTotalPayments.Print TotalPayments
End Sub

Private Sub cmdComputeTax_Click()
    Select Case TaxableIncome                                                           'Case determines what tax bracket the user is in and computes their income tax
        Case Is > 311950
            IncomeTax = 90514.5 + 0.35 * (TaxableIncome - 311950)
                picResultsTax.Print IncomeTax
        Case 143500 To 311950
            IncomeTax = 34926 + 0.33 * (TaxableIncome - 143500)
                picResultsTax.Print IncomeTax
        Case 68800 To 143500
            IncomeTax = 14010 + 0.28 * (TaxableIncome - 68800)
                picResultsTax.Print IncomeTax
        Case 28400 To 68800
            IncomeTax = 3910 + 0.25 * (TaxableIncome - 28400)
                picResultsTax.Print IncomeTax
        Case 7000 To 28400
            IncomeTax = 700 + 0.15 * (TaxableIncome - 7000)
                picResultsTax.Print IncomeTax
        Case 0 To 7000
            IncomeTax = 0.1 * TaxableIncome
                picResultsTax.Print IncomeTax
        Case Else
                picResultsTax.Print "Error"
    End Select
        
            
        
End Sub

Private Sub cmdFedTax_Click()
    Open App.Path & "\Tax.txt" For Input As #2                                  'File is opened
    picResultsFedTax.Cls
    CTR = 0
    Do While Not EOF(2)                                                         'fill the array with data from file
        CTR = CTR + 1
        Input #2, Tax(CTR)
    Loop
    picResultsFedTax.Print Tax(CTR)                                             'The amount of taxes shown on users Form W-2 is displayed
End Sub
