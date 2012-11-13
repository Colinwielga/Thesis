VERSION 5.00
Begin VB.Form frmCashFlow 
   BackColor       =   &H00000000&
   Caption         =   "Cash Flow"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12720
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "Cash Flow.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   12720
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8040
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      DisabledPicture =   "Cash Flow.frx":65AE
      DownPicture     =   "Cash Flow.frx":11772
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cash Flow.frx":1C936
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblEnterName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please enter your first name"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   6000
      Width           =   4215
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Cash Flow Fantasy 2010"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3855
      Left            =   6480
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
End
Attribute VB_Name = "frmCashFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cash Flow Fantasy 2010
'Author: Darcy Delaney
'Date Written: October 25th, 2010
'Program Objective: The following program is a simulation that uses economic principles to illustrate
'                   the pattern of cash flow of select professions within the job market. It presents salaries,
'                   calculates taxes, and allows the user to explore the cost of living in the United States. Figures
'                   used in this program are based on assumptions and taken from several real life databases for the
'                   year 2010. The program allows the user to select from a variety of 18 of the most popular professions,
'                   declares a salary for that particular profession, and subsequently calculate taxes based on that salary.
'                   The user also awarded the option to explore and manipulate the cost of living data extensively. This program
'                   aims to give insight to the finances of working professionals and also illustrate the immediate and fixed expenses
'                   the United States' population at large are faced with.

'Form Objective:    This form is the opening form of the program. It possesses the title, asks for the user's first name and then
'                   to the next form that formally starts the program..
                    
Option Explicit

'Obtain user's first name and directs thenm to the start page
Private Sub cmdEnter_Click()
    frmStart.Visible = True
    frmCashFlow.Visible = False
    UserName = txtUserName.Text
End Sub

