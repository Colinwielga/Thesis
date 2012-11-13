VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10545
   ClientLeft      =   1080
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   Picture         =   "frmHomeForm.frx":0000
   ScaleHeight     =   10545
   ScaleMode       =   0  'User
   ScaleWidth      =   12450
   Begin VB.CommandButton cmdFutureValue 
      BackColor       =   &H0000FF00&
      Caption         =   "Future Value of Investments"
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
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   9000
      Picture         =   "frmHomeForm.frx":15CD
      ScaleHeight     =   4305
      ScaleWidth      =   3345
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin VB.PictureBox picResults1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      Picture         =   "frmHomeForm.frx":549E
      ScaleHeight     =   4305
      ScaleWidth      =   3345
      TabIndex        =   5
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton cdmCapitalBudgeting 
      BackColor       =   &H0000FF00&
      Caption         =   "Capital Budgeting Calculators"
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
      Left            =   1080
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdCompanyInfo 
      BackColor       =   &H0000FF00&
      Caption         =   "Various Companies and Related Financial DataInformation"
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
      Left            =   9480
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdReconciliation 
      BackColor       =   &H0000FF00&
      Caption         =   "Bank Reconciliation"
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
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FF00&
      Caption         =   "Quit Project"
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
      Left            =   10080
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9240
      Width           =   2175
   End
   Begin VB.CommandButton cmdAmort 
      BackColor       =   &H0000FF00&
      Caption         =   "Amortization Calculator"
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
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Financial Instruments and Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3360
      TabIndex        =   10
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Advanced Financial Concepts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Personal Finance Calulators"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   8
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Various Companies and Financial Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   8760
      TabIndex        =   7
      Top             =   4800
      Width           =   3375
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Project1(Financila_Instruments.vbp)
'Form: frmMainMenu(frmHomePage.frm)
'Author: Sean Mase and David Horn
'Date Written: March 26, 2008
'Objective: The purpose of this form is connect all the forms to a
            'single home page.  From this form the user can access all
            ' ohter forms within the project.
            
            'The purpose of this project is to provide users with various
            'financial information that they may find both educational and
            'useful in their daily lives.
            
Option Explicit

Private Sub cdmCapitalBudgeting_Click()
    'displays capital budgeting form
    frmCapitalBudgeting.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdAmort_Click()
    'displays amortization form
    frmAmortization.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdCompanyInfo_Click()
    'displays various company information form
    frmCompanyInfo.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdFutureValue_Click()
    'displays the future value of investments form
    frmFutureValue.Show
    frmMainMenu.Hide
    
End Sub

Private Sub cmdQuit_Click()
    'allows user to quit the program
    End
End Sub

Private Sub cmdReconciliation_Click()
    'displays the bank account reconciliation form
    frmReconciliation.Show
    frmMainMenu.Hide
    
End Sub
