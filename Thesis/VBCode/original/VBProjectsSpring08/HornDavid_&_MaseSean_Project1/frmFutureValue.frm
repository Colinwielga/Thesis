VERSION 5.00
Begin VB.Form frmFutureValue 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExplanation 
      BackColor       =   &H0000FF00&
      Caption         =   "What does this calcluation show?"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Main Menu"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H0000FF00&
      Caption         =   "Calculate Future Value"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
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
      Left            =   6240
      ScaleHeight     =   435
      ScaleWidth      =   1995
      TabIndex        =   11
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox txtCompound 
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
      Left            =   6240
      TabIndex        =   8
      Text            =   "12"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtInterest 
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
      Left            =   6240
      TabIndex        =   3
      Text            =   ".10"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtYears 
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
      Left            =   6240
      TabIndex        =   2
      Text            =   "30"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtContribution 
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
      Left            =   6240
      TabIndex        =   1
      Text            =   "1500"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtPrincipal 
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
      Left            =   6240
      TabIndex        =   0
      Text            =   "10000"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Future Value of Investment (Retirement Planning)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Future Value is"
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
      Left            =   3360
      TabIndex        =   12
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "times per year."
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
      TabIndex        =   10
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compound interest"
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
      Left            =   2760
      TabIndex        =   9
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Expected Annual Return (in decimal form)."
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
      Left            =   2160
      TabIndex        =   7
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Years to Retirement"
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
      Left            =   3720
      TabIndex        =   6
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contribution Per Compound Period"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current  Principal"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
End
Attribute VB_Name = "frmFutureValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Project1(Financila_Instruments.vbp)
'Form: frmFutureValue(frmFutureValue.frm)
'Author: Sean Mase and David Horn
'Date Written: March 27, 2008
'Objective: the purpose of this form is to provide users the ability to see the amount of money an investment will grow
            'to in the futures.  The importance of this is that the user is able to better plan for retirement by seeing
            'what there current saving patterns will finally result in. Also, this form shows the power of compound interest.

Option Explicit


Private Sub cmdCalculate_Click()
    'declares variables
    Dim Principal As Double, Contribution As Double, Years As Integer, Interest As Single, Total As Double
    Dim Compound As Integer, n As Integer, i As Single, FVPrincipal As Double, FVAnnuity As Double

    'assign values to variables through user input
    Principal = txtPrincipal.Text
    Contribution = txtContribution.Text
    Years = txtYears.Text
    Interest = txtInterest.Text
    Compound = txtCompound.Text
    
    'the purpose of the variables below is to make the future value calculations more simple to input
    n = Years * Compound
    i = Interest / Compound
    
    'Future value calculations
    FVPrincipal = Principal * (1 + i) ^ n
    FVAnnuity = Contribution * (((1 + i) ^ n - 1) / i)
    Total = FVPrincipal + FVAnnuity
    
    'Prints the results
    picResults.Cls
    picResults.Print FormatCurrency(Total, 2)
    
End Sub

Private Sub cmdExplanation_Click()
    'displays message explaining the purpose of the future value calculation
    MsgBox ("The purpose of this future value calculation is two fold. 1) It helps users plan for retirement" _
        & " by calculating the future value of the users saving habits. 2) This clalculation helps demonstrate the power" _
        & " of compound interest.  To demonstrate compound interest change the years to retirement from 30 to 20 years. " _
        & "You will witness a drastic decrease of the future value of your investments.  This decrease is due to the fact " _
        & "your investment will not be earning interest on the interest that is earned in the 10 years lost. Overall, the" _
        & " the moral of the story is you should start saving for retirement ASAP.")
End Sub

Private Sub cmdMainMenu_Click()
    'displays main menu
    frmMainMenu.Show
    frmFutureValue.Hide
    
End Sub


