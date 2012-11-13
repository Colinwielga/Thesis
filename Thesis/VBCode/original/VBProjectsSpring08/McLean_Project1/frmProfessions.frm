VERSION 5.00
Begin VB.Form frmProfessions 
   Caption         =   "Professions"
   ClientHeight    =   11760
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   14400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11760
   ScaleWidth      =   14400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContents 
      Caption         =   "Back to the Table of Contents"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   7920
      TabIndex        =   7
      Top             =   9840
      Width           =   2892
   End
   Begin VB.CommandButton cmdCFO 
      Caption         =   "Chief Financial Officer"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   1200
      TabIndex        =   6
      Top             =   9840
      Width           =   2892
   End
   Begin VB.CommandButton cmdController 
      Caption         =   "Corporate Controller"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   1200
      TabIndex        =   5
      Top             =   6240
      Width           =   2892
   End
   Begin VB.CommandButton cmdTreasurer 
      Caption         =   "Corporate Treasurer"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   1200
      TabIndex        =   4
      Top             =   8040
      Width           =   2892
   End
   Begin VB.CommandButton cmdFinAnalyst 
      Caption         =   "Financial Analyst"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   1200
      TabIndex        =   3
      Top             =   4440
      Width           =   2892
   End
   Begin VB.CommandButton cmdTax 
      Caption         =   "Tax Accountant"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   2892
   End
   Begin VB.CommandButton cmdAudit 
      Caption         =   "Audit Accountant"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2892
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   16.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6132
      Left            =   5520
      ScaleHeight     =   6084
      ScaleWidth      =   7524
      TabIndex        =   0
      Top             =   2880
      Width           =   7572
   End
   Begin VB.Image Image1 
      Height          =   11772
      Left            =   0
      Picture         =   "frmProfessions.frx":0000
      Top             =   0
      Width           =   14400
   End
End
Attribute VB_Name = "frmProfessions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Accounting Project
'Profession Form
'Tony McLean
'3.31.2008
'The purpose of this form to allow the user to learn about
'various careers in the field of accounting by reading
'various professional profiles.
Option Explicit
'Displays textual information pertaining to the profession the user would like to read about
Private Sub cmdAudit_Click()
    picResults.Cls
    picResults.Print "The Staff auditor performs the detail"
    picResults.Print "work of a financial audit under the"
    picResults.Print "supervision of a Senior. Staff Auditors"
    picResults.Print "will often start to direct small audits"
    picResults.Print "at the two-year level."
    picResults.Print ""
    picResults.Print "The Senior auditor works under the general"
    picResults.Print "direction of an Audit Manager. Responsibilities"
    picResults.Print "include the direction of audit field work,"
    picResults.Print "assignment of detail work to Staff, and review"
    picResults.Print "of thier working papers.  The Senior auditor also"
    picResults.Print "prepares financial statements, develops corporate"
    picResults.Print "tax returns, and suggests improvements to"
    picResults.Print "internal controls."
End Sub
'Displays textual information pertaining to the profession the user would like to read about
Private Sub cmdCFO_Click()
    picResults.Cls
    picResults.Print "The Chief Financial Officer (CFO) is typically"
    picResults.Print "designated Vice President of Finance. The CFO"
    picResults.Print "advises the President of the organization"
    picResults.Print "with respect to financial reporting, financial"
    picResults.Print "stability and liquidity, and financial growth."
    picResults.Print "The CFO Directs and supervises the work of"
    picResults.Print "the Controller, Treasurer, and sometimes the"
    picResults.Print "Internal Auditing Manager. Other duties may"
    picResults.Print "include maintenance of relationships with"
    picResults.Print "stockholders, financial institutions, and the"
    picResults.Print "investment community. Frequently, the CFO is a"
    picResults.Print "member of the Board of Directors and/or the"
    picResults.Print "Executive Committee and as such, contributes"
    picResults.Print "to overall organization planning, policy"
    picResults.Print "development, and implementation."
End Sub
'Displays textual information pertaining to the profession the user would like to read about
Private Sub cmdContents_Click()
    frmProfessions.Hide
    frmFirms.Hide
    frmSalaries.Hide
    frmDidKnow.Hide
    frmContents.Show
    frmIntroduction.Hide
End Sub
'Displays textual information pertaining to the profession the user would like to read about
Private Sub cmdController_Click()
    picResults.Cls
    picResults.Print "The Controller functions as the Chief"
    picResults.Print "Accounting Executive responsible for"
    picResults.Print "organizing, directing, and controlling"
    picResults.Print "the work of the accounting personnel in"
    picResults.Print "collecting, summarizing, and interpreting"
    picResults.Print "financial data for the use of management,"
    picResults.Print "creditors, investors, and taxing authorities."
    picResults.Print "As a member of the top management team, the"
    picResults.Print "Controller helps develop forecasts for"
    picResults.Print "proposed projects of the organization,"
    picResults.Print "measures actual performance against operating"
    picResults.Print "plans and standards, and interprets the"
    picResults.Print "results of operations for all levels of"
    picResults.Print "management."
End Sub
'Displays textual information pertaining to the profession the user would like to read about
Private Sub cmdFinAnalyst_Click()
    picResults.Cls
    picResults.Print "A Financial Analyst works at the direction"
    picResults.Print "of a Senior or Manager in performing various"
    picResults.Print "financial or budget analyses. Assignments"
    picResults.Print "are in one area or several, including profit"
    picResults.Print "planning, capital expenditures, investments,"
    picResults.Print "cash flow budgeting, and acquisitions."
End Sub
'Displays textual information pertaining to the profession the user would like to read about
Private Sub cmdTax_Click()
    picResults.Cls
    picResults.Print "The Tax Accountant prepares tax returns,"
    picResults.Print "researches tax questions, and counsels clients "
    picResults.Print "on tax problems under the supervision of a"
    picResults.Print "Tax Senior and/or Tax Manager."
    picResults.Print ""
    picResults.Print "The Tax Senior works under the general"
    picResults.Print "direction of a Tax Manager and/or Tax"
    picResults.Print "Partner. This accountant prepares or reviews"
    picResults.Print "tax returns for individuals and organizations,"
    picResults.Print "researches tax questions, offers suggestions for"
    picResults.Print "tax planning, and studies law for potential"
    picResults.Print "tax savings."
End Sub
'Displays textual information pertaining to the profession the user would like to read about
Private Sub cmdTreasurer_Click()
    picResults.Cls
    picResults.Print "The Treasurer directs the functions dealing"
    picResults.Print "largely with the receipt, disbursement, and"
    picResults.Print "protection of cash, the preservation of company"
    picResults.Print "assets, and the investment of surplus funds or"
    picResults.Print "pension and trust funds.  The Treasurer also"
    picResults.Print "determines the optimal cash position for the"
    picResults.Print "organization and sets short-term investment"
    picResults.Print "policies.  The Treasurer Governs overall credit"
    picResults.Print "policy, negotiates loans, arranges insurance"
    picResults.Print "coverage, and maintains banking relationships."
End Sub
