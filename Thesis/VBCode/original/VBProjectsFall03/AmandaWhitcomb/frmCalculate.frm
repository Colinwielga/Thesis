VERSION 5.00
Begin VB.Form frmCalculate 
   BackColor       =   &H000080FF&
   Caption         =   "Calculate"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Enabled         =   0   'False
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5640
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H000080FF&
      Caption         =   "Calculate the Total Cost of your trip"
      Height          =   975
      Left            =   3240
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.PictureBox results 
      BackColor       =   &H0080C0FF&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1515
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Image imgpalmtree 
      Height          =   4695
      Left            =   5280
      Picture         =   "frmCalculate.frx":0000
      Top             =   120
      Width           =   3450
   End
End
Attribute VB_Name = "frmCalculate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Exotica_Travel (Amanda Whitcomb's VBProject.vbp)
'Form Name : frmCalculate (frmCalculate.frm)
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

Private Sub cmdCalculate_Click()
'Clear whatever may be in Results for repeated use.
results.Cls

'declare all variables
totalcost = ff + hh + ee + (totalmeal + (totalmeal * 0.2))

'Print out the headings preceding total cost for the users trip.
results.Print "             You have finished making your travel arrangements."
results.Print "      "
results.Print "~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~..~"
results.Print "      "

'determine total cost, decide remaining
'trip money, print results
If totalcost > 6000 Then
    results.Print "             You have exceeded your trip limit and will"
    results.Print "             need to make new selections or pay the difference."
    results.Print "             The differece is "; FormatCurrency(totalcost - 6000)
ElseIf totalcost = 6000 Then
    results.Print "You have met your trip limit!"
ElseIf totalcost < 6000 Then
    results.Print "             WOW! You still have trip money! The remaining "
    results.Print "            "; FormatCurrency(6000 - totalcost); " is yours for spending."
End If

'Disable Calculate and enable
'restart button and quit button
cmdCalculate.Enabled = False
cmdRestart.Enabled = True
cmdQuit.Enabled = True
End Sub

Private Sub cmdQuit_Click()
'enable calculate for repeated use
cmdCalculate.Enabled = True
'clear results for repeated use
results.Cls
'end program
    End
End Sub

Private Sub cmdRestart_Click()
'enable calculate for repeated use
cmdCalculate.Enabled = True
'clear results
results.Cls
'Hide Calculate form and show
'travel destination form for
'program restart
    frmCalculate.Hide
    FrmTravelDestination.Show
End Sub
