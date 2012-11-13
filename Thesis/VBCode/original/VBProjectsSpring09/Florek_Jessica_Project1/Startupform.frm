VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   Picture         =   "Startupform.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H000000FF&
      Caption         =   "Start Planning"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H000000FF&
      Caption         =   "Continue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmForm1
'Jessica Florek
'Written: 3/4/09
'Objective: First slide of project, Gets input from user such as their budget
'and duration of trip.



Option Explicit

Private Sub cmdContinue_Click()

'calculates costs for use later in program
foodcost = duration * 25
budget = budget - foodcost

frmStart.Hide
frmAirplane.Show

'sets up each city as a false boolean so that they summary of each cities expenses will only be displayed if the city is clicked on, therefore becoming true
venice = False
paris = False
madrid = False
london = False

End Sub

Private Sub cmdQuit_Click()
End
End Sub




Private Sub cmdStart_Click()
'gathers necessary information needed for the program such as the budget and duration of the trip
budget = InputBox("Enter amount budgeted for your trip.")
budget2 = budget
euros = budget / 1.26
MsgBox ("You have " & FormatCurrency(Round(euros), 0) & ", in euros, for your trip!")
duration = InputBox("Enter how many days you plan to travel.")
cmdContinue.Enabled = True
cmdStart.Enabled = False

End Sub


