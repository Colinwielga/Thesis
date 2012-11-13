VERSION 5.00
Begin VB.Form frmClaimed 
   BackColor       =   &H00000000&
   Caption         =   "Claimed Dependent "
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H000000FF&
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   3120
      ScaleHeight     =   2475
      ScaleWidth      =   5355
      TabIndex        =   9
      Top             =   3960
      Width           =   5415
   End
   Begin VB.CommandButton cmdPmt 
      BackColor       =   &H000000FF&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H000000FF&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdPotentialAudit 
      BackColor       =   &H000000FF&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
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
      Top             =   5400
      Width           =   1695
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
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
      Caption         =   "3.) Enter Federal Income Taxes Withheld (i.e. W-2)"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3240
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00000000&
      Caption         =   "2.) Compute your Federal Income Tax"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00000000&
      Caption         =   "1.) Determine Potential for an audit. (Scale 1-6)"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmClaimed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
     'Determines the amount of income subject to income tax

    TaxableIncome = AGI - 3300
    
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
    frmClaimed.Hide
    frmSave.Show
    
End Sub
