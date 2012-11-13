VERSION 5.00
Begin VB.Form frmSingle 
   BackColor       =   &H0000FFFF&
   Caption         =   "Single"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FF00&
      Caption         =   "Save "
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdDeduct 
      BackColor       =   &H0000FF00&
      Caption         =   "Determine Deductions"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdPmt 
      BackColor       =   &H0000FF00&
      Caption         =   "Refund or Payment?"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdPotentialAudit 
      BackColor       =   &H0000FF00&
      Caption         =   "Potential Audit"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FF00&
      Caption         =   "Exit"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H0000FF00&
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   2640
      ScaleHeight     =   2475
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   4560
      Width           =   5415
   End
   Begin VB.Label lbl5 
      BackColor       =   &H0000FFFF&
      Caption         =   "3.) Compute your Federal Income Tax"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   10
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lbl3 
      BackColor       =   &H0000FFFF&
      Caption         =   "2.) Determine Potential for an audit. (Scale 1-6)"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lbl2 
      BackColor       =   &H0000FFFF&
      Caption         =   "1.) Itemized or Standard Deductions?"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lbl1 
      BackColor       =   &H0000FFFF&
      Caption         =   "4.) Enter Federal Income Taxes Withheld (i.e. W-2)"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label lbl4 
      BackColor       =   &H0000FFFF&
      Caption         =   "How to Begin?"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Financial Information......."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   2655
   End
End
Attribute VB_Name = "frmSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCompute_Click()
    'Chooses the better of standardized or Itemized
    Standard = 5150
    If Itemized > Standard Then
        Deduct = Itemized
    Else
        Deduct = Standard
    End If
    'Determines the amount of income subject to income tax
    TaxableIncome = AGI - 3300 - Deduct
    
    'Determining the percentage bracket
    Select Case TaxableIncome
        Case Is >= 336550
            TaxLiability = 97653 + (TaxableIncome - 336550) * 0.35
            MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 154800 To 336550
            TaxLiability = 37675.5 + (TaxableIncome - 154800) * 0.33
             MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 74200 To 154800
            TaxLiability = 15107.5 + (TaxableIncome - 74200) * 0.28
             MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 30650 To 74200
            TaxLiability = 4220 + (TaxableIncome - 30650) * 0.25
             MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 7550 To 30650
            TaxLiability = 755 + (TaxableIncome - 7550) * 0.15
             MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 0 To 7550
            TaxLiability = TaxableIncome * 0.1
            MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case Is < 0
            MsgBox "You have no tax liability ", , "Results"
            TaxLiability = 0
        Case Else
            MsgBox "Sorry. You have entered an invalid Taxable income", , "Error"
    End Select
    
    picResults.Print "***************************"
    picResults.Print "Adjusted Gross Income: "; FormatCurrency(AGI)
    picResults.Print "Taxable Income: "; FormatCurrency(TaxableIncome)
    picResults.Print "Tax Liability: "; FormatCurrency(TaxLiability)
    picResults.Print "***************************"
    
End Sub

Private Sub cmdDeduct_Click()
    frmSingle.Hide
    frmDeduct.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPmt_Click()
    'Requires another input by user
    TaxesWithheld = InputBox("Please enter amount of taxes withheld i.e Obtain from W-2", "Input")
    
    'Determines refund or paymentIf TaxesWithheld > TaxLiability Then
    If TaxesWithheld > TaxLiability Then
        Refund = (TaxesWithheld - TaxLiability)
        picResults.Print "Your refund is  "; FormatCurrency(Refund)
    Else
        Payment = TaxLiability - TaxesWithheld
        picResults.Print "You owe  "; FormatCurrency(Payment)
        
    End If
    
    
End Sub

Private Sub cmdPotentialAudit_Click()
    
    AGI = InputBox("Please enter your Adjusted Gross Income($)", "Input")
    'This will perform a match & stop search to display potential audit
    Pos = 0
    Found = False
    
    Do Until Found = True Or Pos > CTR1
        Pos = Pos + 1
        If AGI > Bracket(Pos) Then
            Found = True
        End If
    Loop
    
    If Found = True Then
        picResults.Print "Name: "; N
        picResults.Print "***************************"
        picResults.Print Risk(Pos); ":"; Potential(Pos)
    Else
        MsgBox " Check your AGI input!", , "Error"
    End If
    
End Sub

Private Sub cmdSave_Click()
    frmSingle.Hide
    frmSave.Show
End Sub

