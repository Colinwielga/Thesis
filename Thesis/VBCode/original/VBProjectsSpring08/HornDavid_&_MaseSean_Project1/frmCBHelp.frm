VERSION 5.00
Begin VB.Form frmCBHelp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTV 
      BackColor       =   &H0000FF00&
      Caption         =   "Terminal Value"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdRV 
      BackColor       =   &H0000FF00&
      Caption         =   "Residual Value"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdCS 
      BackColor       =   &H0000FF00&
      Caption         =   "Annual Cost Savings"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdCashFlows 
      BackColor       =   &H0000FF00&
      Caption         =   "Annual Cash Flows"
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
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdNPV 
      BackColor       =   &H0000FF00&
      Caption         =   "Net Present Value"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdPP 
      BackColor       =   &H0000FF00&
      Caption         =   "Payback Period"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdMPP 
      BackColor       =   &H0000FF00&
      Caption         =   "Minimum Payback Period"
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
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdMTR 
      BackColor       =   &H0000FF00&
      Caption         =   "Marginal Tax Rate"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdCOC 
      BackColor       =   &H0000FF00&
      Caption         =   "Cost of Capital"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdUL 
      BackColor       =   &H0000FF00&
      Caption         =   "Useful Life"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdII 
      BackColor       =   &H0000FF00&
      Caption         =   "Initial Investment"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdCB 
      BackColor       =   &H0000FF00&
      Caption         =   "Capital Budgeting"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Capital Budgeting  Form"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Capital Budgeting Definitions"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   360
      TabIndex        =   13
      Top             =   360
      Width           =   8055
   End
End
Attribute VB_Name = "frmCBHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Project1(Financila_Instruments.vbp)
'Form: frmCBHelp(frmCBHelp)
'Author: Sean Mase and David Horn
'Date Written: March 26, 2008
'Objective: The purpose of this form is to provide users with
            'definitions of terms that are used on the CapitalBudgeting
            'form.  This will allow users to better understand
            'what information to input into the various text boxes.

Option Explicit

Private Sub cmdCashFlows_Click()
    'displays definition for annual cash flows
    MsgBox ("Annual Cash Flow is the annual cash intake that is expected to be realized " _
        & "if a capital budgeting project is accepted. These cash flows can" _
        & " derive from increases in company sales by accepting the new project or " _
        & " costs savings resulting in increased cash flows.  For sake of simplicity, " _
        & "this project assumes that annual cash flows are constant from year to year. " _
        & "Note that this is often not the case.")
End Sub

Private Sub cmdCB_Click()
    'diplays definition for capital budgeting
    MsgBox ("Capital Budgeting is the decision making process used by" _
        & " managers when they are deciding whether to accept a new project.")
End Sub

Private Sub cmdCOC_Click()
    'displays definition for cost of capital
    MsgBox ("The cost of capital is the required return a project must generate in order for a" _
        & " company to accept it.")
End Sub

Private Sub cmdCS_Click()
    'displays definition for annual cost savings
    MsgBox ("See annual cash flows.")
End Sub

Private Sub cmdII_Click()
    'displays definition for initial investment
    MsgBox ("The initial investment (ii) in a new project is all the cash inflows" _
        & " aand outflows that take place as soon as the investment begins." _
        & "An examples of ii cash inflow are the sale of old equipment the new project" _
        & " is replacing.  An example of ii cash outflow is the cost of the new project" _
        & "itself.")
    
End Sub

Private Sub cmdMPP_Click()
    'diplays definition for minimum payback period
    MsgBox ("The minimum payback period is the length of time, normally years, a company" _
        & " requires for a projects initial investment to be recovered.")
End Sub

Private Sub cmdMTR_Click()
    'displays definition for marginal tax rate
    MsgBox ("Since the US tax code uses a progressixe tax rate, a company or individual" _
        & "tax rate increases as there income level rises.  A company's marginal tax " _
        & "is the tax rate a company will incur for increasing there income.")
End Sub

Private Sub cmdNPV_Click()
    'displays definition for net present value
    MsgBox ("Net Present Value (NPV) compares the value of a dollar today to the value of " _
        & "that same dollar in the future, taking inflation and interest into account. If " _
        & "the NPV of an investment is positive, it should be accepted. IF NPV is negative," _
        & " the project should be rejected.")
End Sub

Private Sub cmdPP_Click()
        'displays definition for payback period
        MsgBox ("A project's payback period is the length of time, normally years, it " _
        & " takes to recover its initial investment.")
End Sub

Private Sub cmdReturn_Click()
    'displays capital budgeting form
    frmCapitalBudgeting.Show
    frmCBHelp.Hide
End Sub

Private Sub cmdRV_Click()
    'displays definition for residual value
    MsgBox ("Residual value is the value of the project when it is sold at the end of its" _
        & "  usefule life. Also known as terminal value, salvage value, " _
        & "and scrap value.")
End Sub

Private Sub cmdTV_Click()
    'displays definition for terminal value
    MsgBox ("See residual Value.")
End Sub

Private Sub cmdUL_Click()
    'displays definition for useful life
    MsgBox ("The useful life of an investment is the time, normally in years, the" _
        & " new investment is expected to last.")
End Sub


