VERSION 5.00
Begin VB.Form frmBegin 
   BackColor       =   &H008080FF&
   Caption         =   "Who am I?"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter and Save Personal Information"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0FF&
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdRate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Review This Program"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdPattern 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pattern Test"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      MaskColor       =   &H80000004&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdSocialPsych 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Social Psych Characteristics"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "By: Sara Peplinski, C-SCI 130"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   8040
      Width           =   3135
   End
   Begin VB.Label lblTitleDescription 
      BackColor       =   &H008080FF&
      Caption         =   $"frmBegin.frx":0000
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   6735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Who am I?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sara Peplinski, C-SCI 130
'Project Name: Who am I in Social Psychology
'Date Written: March 17, 2006
'This program overall gives the user a brief look at some concepts in social psychology.
'It is also provides a fun personality test for the user to complete about himself or herself.
'The user should end the program with the review section in order to show comprehension of
'the material found in the program.


Private Sub cmdEnd_Click()
'end of program
    End
End Sub
Private Sub cmdPattern_Click()
'command button that transfers user to Pattern Test
    MsgBox "Here is a fun, short personality test!", , "Message"
    frmBegin.Hide
    frmPatternTest.Show
End Sub
Private Sub cmdRate_Click()
'command button that transfers user to Review of the Program
    frmBegin.Hide
    frmReview.Show
End Sub
Private Sub cmdSave_Click()
'Enter personal information about the user
    
    'these variables are declared in the module as Public in order to use
    'this information in the Social Psychology Section
    First = InputBox("Enter First Name", "First Name")
    Last = InputBox("Enter Last Name", "Last Name")
    MsgBox "Now Enter Birthday Information", , "Birthday"
    BMonth = InputBox("Month [MM]", "Month")
    BDay = InputBox("Day [DD]", "Day")
    BYear = InputBox("Year [YY]", "Year")
End Sub
Private Sub cmdSocialPsych_Click()
'command button that transfers user to Social Psychology review area
    MsgBox "This area has two short activities as well as reviews theories about the self", , "Message"
    frmBegin.Hide
    frmSocial.Show
End Sub

