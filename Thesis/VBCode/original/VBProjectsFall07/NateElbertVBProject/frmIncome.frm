VERSION 5.00
Begin VB.Form frmIncome 
   Caption         =   "Income Section of Tax Return"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20265
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   20265
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResultWages 
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
      Left            =   17160
      ScaleHeight     =   555
      ScaleWidth      =   2835
      TabIndex        =   35
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdLoad 
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
      Left            =   13920
      TabIndex        =   34
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit2 
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
      Height          =   1095
      Left            =   12840
      TabIndex        =   33
      Top             =   6840
      Width           =   3975
   End
   Begin VB.CommandButton cmdGoBack 
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
      Height          =   1095
      Left            =   8160
      TabIndex        =   32
      Top             =   6840
      Width           =   4095
   End
   Begin VB.CommandButton cmdContinue2 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   31
      Top             =   6840
      Width           =   4575
   End
   Begin VB.PictureBox picResultsTaxableIncome 
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
      Left            =   17280
      ScaleHeight     =   555
      ScaleWidth      =   2715
      TabIndex        =   30
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton cmdComputeTaxableIncome 
      Caption         =   "Compute Taxable Income"
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
      Left            =   13920
      TabIndex        =   29
      Top             =   6000
      Width           =   3015
   End
   Begin VB.PictureBox picResultsNo 
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
      Left            =   13920
      ScaleHeight     =   555
      ScaleWidth      =   2955
      TabIndex        =   28
      Top             =   4080
      Width           =   3015
   End
   Begin VB.PictureBox picResultsAGI 
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
      Left            =   17280
      ScaleHeight     =   555
      ScaleWidth      =   2715
      TabIndex        =   27
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdComputeAGI 
      Caption         =   "Compute AGI"
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
      Left            =   13920
      TabIndex        =   26
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txtUnemploymentCompensation 
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
      Left            =   13920
      TabIndex        =   25
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtTaxableInterest 
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
      Left            =   13920
      TabIndex        =   24
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CheckBox chkNo 
      Caption         =   "No"
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
      Left            =   4320
      TabIndex        =   16
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox chkYes 
      Caption         =   "Yes"
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
      Left            =   3000
      TabIndex        =   15
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lbl6b 
      Caption         =   "6"
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
      Left            =   12840
      TabIndex        =   23
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label lbl5b 
      Caption         =   "5"
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
      Left            =   12840
      TabIndex        =   22
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lbl4b 
      Caption         =   "4"
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
      Left            =   12840
      TabIndex        =   21
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label lbl3b 
      Caption         =   "3"
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
      Left            =   12840
      TabIndex        =   20
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lbl2b 
      Caption         =   "2"
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
      Left            =   12840
      TabIndex        =   19
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lbl1b 
      Caption         =   "1"
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
      Left            =   12840
      TabIndex        =   18
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblNote3 
      Caption         =   "Note. You MUST check Yes or No."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   17
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lbl6a 
      Caption         =   "6"
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
      TabIndex        =   14
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label lbl5a 
      Caption         =   "5"
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
      TabIndex        =   13
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lbl4a 
      Caption         =   "4"
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
      TabIndex        =   12
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label lbl3a 
      Caption         =   "3"
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
      TabIndex        =   11
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lbl2a 
      Caption         =   "2"
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
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblNote2 
      Caption         =   "Enclose, but do not attach, any payment."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblNote1 
      Caption         =   "Attach form W-2 here."
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
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblIncome 
      Caption         =   "Income"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lbl1a 
      Caption         =   "1"
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
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblTaxableIncome 
      Caption         =   "Subtract line 5 from line 4. If line 5 is larger than line 4, enter -0-."
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
      Left            =   3000
      TabIndex        =   5
      Top             =   6000
      Width           =   8055
   End
   Begin VB.Label lblClaim 
      Caption         =   "Can your parents (or someone else) claim you on their return?"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   4080
      Width           =   8055
   End
   Begin VB.Label lblAGI 
      Caption         =   "Add lines 1, 2, and 3. This is your Adjusted Gross Income."
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
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   8055
   End
   Begin VB.Label lblUnemployment 
      Caption         =   "Unemployment compensation and Alaska Permanent Fund dividends."
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
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Width           =   8055
   End
   Begin VB.Label lblInterest 
      Caption         =   "Taxable Interest. If the total is over $1,500, you cannot use Form 1040EZ."
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
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   8055
   End
   Begin VB.Label lblWages 
      Caption         =   "Wages, salaries, and tips. This should be in box 1 of Form W-2."
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
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frmIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkNo_Click()
        picResultsNo.Cls
    If chkNo.Value = 1 Then                                                            'If checked the amount of deductions is displayed
        picResultsNo.Print Int(7800)
    End If
    
End Sub

Private Sub chkYes_Click()
    If chkYes.Value = 1 Then                                                            ' If checked the program displays an error message
        MsgBox "You should not be using this Form to file your Taxes, please exit the program and consult your accountant!"
    End If
End Sub

Private Sub cmdComputeAGI_Click()
                                                                                        'Read components of AGI from textbox
    TaxableInterest = txtTaxableInterest.Text
    UnemploymentCompensation = txtUnemploymentCompensation.Text
                    picResultsAGI.Cls
            If UnemploymentCompensation < 1 Then                                        'Compute AGI and display
                AGI = Wages(CTR) + TaxableInterest
                    picResultsAGI.Print AGI
            Else
                AGI = Wages(CTR) + TaxableInterest + UnemploymentCompensation           'Compute AGI and display
                    picResultsAGI.Print AGI
            End If
            
End Sub

Private Sub cmdComputeTaxableIncome_Click()
    NoClaim = 7800
        If NoClaim > AGI Then                                                               'If line 5 is larger then line 4, then a zero is entered
            picResultsTaxableIncome.Print 0
                TaxableIncome = 0
            Else
                TaxableIncome = AGI - NoClaim                                                   'Computation of Taxable Income
                    picResultsTaxableIncome.Print TaxableIncome
        End If
    
End Sub


Private Sub cmdContinue2_Click()
    frmPaymentsAndTax.Show                                                              'Button hides the Income form and shows the Payments and Tax form
    frmIncome.Hide
End Sub

Private Sub cmdGoBack_Click()
    frmLabel.Show                                                                       'Button returns the user to the previous page
    frmIncome.Hide
End Sub

Private Sub cmdLoad_Click()
    Open App.Path & "\Wages.txt" For Input As #1                                         'Opens file to read wages earned for the year
        picResultWages.Cls
        CTR = 0
            Do While Not EOF(1)                                                                 'fill the array with data from file
                CTR = CTR + 1
                    Input #1, Wages(CTR)
            Loop
        picResultWages.Print Wages(CTR)                                                     'Wages is displayed
    
    
End Sub

Private Sub cmdQuit2_Click()
    End                                                                                     'Button ends program
End Sub

