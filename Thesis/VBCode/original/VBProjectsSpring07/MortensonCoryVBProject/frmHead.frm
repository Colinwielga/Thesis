VERSION 5.00
Begin VB.Form frmHead 
   BackColor       =   &H00000000&
   Caption         =   "Head"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080FFFF&
      Caption         =   "Save"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   2520
      ScaleHeight     =   2475
      ScaleWidth      =   5355
      TabIndex        =   10
      Top             =   4440
      Width           =   5415
   End
   Begin VB.CommandButton cmdPmt 
      BackColor       =   &H0080FFFF&
      Caption         =   "Refund or Payment?"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H0080FFFF&
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdPotentialAudit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Potential Audit"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdDeduct 
      BackColor       =   &H0080FFFF&
      Caption         =   "Determine Deductions"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
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
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
     
    Exemptions = InputBox("Please enter total amount of exemptions, including eligible dependents", "Input")
    
    Standard = 7550
    
    'Chooses the better of standardized or Itemized
    If Itemized > Standard Then
        Deduct = Itemized
    Else
        Deduct = Standard
    End If
    
    'Determines the amount of income subject to income tax
    Exemptions = Exemptions * 3300
    TaxableIncome = AGI - Exemptions - Deduct
    
    'Determining the percentage bracket
    Select Case TaxableIncome
        Case Is >= 336550
            TaxLiability = 94656.5 + (TaxableIncome - 336550) * 0.35
            MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 171650 To 336550
            TaxLiability = 40239.5 + (TaxableIncome - 171650) * 0.33
             MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 106000 To 171650
            TaxLiability = 21857.5 + (TaxableIncome - 106000) * 0.28
             MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 41050 To 106000
            TaxLiability = 5620 + (TaxableIncome - 41050) * 0.25
             MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 10750 To 41050
            TaxLiability = 1075 + (TaxableIncome - 10750) * 0.15
             MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case 0 To 10750
            TaxLiability = TaxableIncome * 0.1
            MsgBox "Your Tax Liability is  " & FormatCurrency(TaxLiability), , "Results"
        Case Is < 0
            MsgBox "You have no tax liability ", , "Results"
            TaxLiability = 0
        Case Else
            MsgBox "Sorry. You have entered an invalid Taxable income", , "Error"
    End Select
    
    Credits = InputBox("Enter dollar amount of tax credits ($)", "Credit")
    Children = InputBox("How many eligible children do you have?", "Child Care Credit")
    TaxLiability = TaxLiability - Credits - (Children * 1000)
    
    picResults.Print "***************************"
    picResults.Print "Adjusted Gross Income: "; FormatCurrency(AGI)
    picResults.Print "Taxable Income: "; FormatCurrency(TaxableIncome)
    picResults.Print "Tax Liability: "; FormatCurrency(TaxLiability)
    picResults.Print "***************************"
End Sub

Private Sub cmdDeduct_Click()
    frmHead.Hide
    frmDeduct.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPmt_Click()
     'Requires another input by user
    TaxesWithheld = InputBox("Please enter amount of taxes withheld i.e Obtain from W-2", "Input")
    
    'Determines refund or payment
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
    frmHead.Hide
    frmSave.Show
End Sub
