VERSION 5.00
Begin VB.Form frmCongratulations 
   BackColor       =   &H00C000C0&
   Caption         =   "Congratulations"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   855
      Left            =   5760
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF80FF&
      Caption         =   $"frmCongratulations.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Image imgcongrats 
      Height          =   930
      Left            =   120
      Picture         =   "frmCongratulations.frx":00EC
      Top             =   360
      Width           =   5085
   End
End
Attribute VB_Name = "frmCongratulations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Exotica_Travel (Amanda Whitcomb's VBProject.vbp)
'Form Name : frmCongratulations (frmCongratulations.frm)
'Author: Amanda Whitcomb
'Date Written: October 30th, 2003
'Purpose of Form:   The user wins an exotic trip
                    'worth $6000 and is given opportunity
                    'to make specialized travel plans. The
                    'program will determine if the user has
                    'overspent or underspent their winnings.
Option Explicit
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Private Sub cmdContinue_Click()
    frmCongratulations.Hide
    FrmTravelDestination.Show
End Sub


