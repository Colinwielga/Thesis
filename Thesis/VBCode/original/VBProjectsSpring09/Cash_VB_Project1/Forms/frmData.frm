VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Scoring Data"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   Picture         =   "frmData.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCourses 
      Caption         =   "Check Out St. Cloud Area Golf Courses"
      Height          =   1095
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   6015
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "Manage Scoring Data"
      Height          =   1095
      Left            =   2400
      TabIndex        =   2
      Top             =   4080
      Width           =   4455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   6480
      TabIndex        =   1
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Golf Guide"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: GolfGuide
':Form Name:  frmTitle
':Author:   Tyler Cash
':Date written:  March 19, 2009

'This program is designed to accomplish three tasks:
'1.) Allow the user to browse St. Cloud Area Golf Courses
'2.) Allow the user to input golf scores and keep a record of those scores.
'3.) Calculate useful statistics from the users golf record.

'**********************************************************************************
'                                 IMPORTANT
'   In order to use this program, you must enable the "Microsoft Scripting Runtime"
'   reference.
'   To do this:
'   1.) Click the project menu
'   2.) Select "References..."
'   3.) Make sure the box next to "Microsoft Scripting Runtime" is checked
'
'**********************************************************************************
'
Option Explicit

Private Sub cmdCourses_Click()
'This button changes to the form allowing the user to select a course to view
    frmCourses.Show
    frmTitle.Hide
End Sub

Private Sub cmdData_Click()
'This button changes to the form allowing the user to manipulate scoring data
    frmStats.Show
    frmTitle.Hide
End Sub

Private Sub cmdQuit_Click()
'This button ends the program
    End
End Sub

