VERSION 5.00
Begin VB.Form FrmFlight 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Show the Flight Cost for your chosen Destination"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6840
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.PictureBox results 
      BackColor       =   &H008080FF&
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   5160
      Width           =   5655
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6840
      TabIndex        =   4
      Top             =   5160
      Width           =   1815
   End
   Begin VB.OptionButton optEconomyClass 
      BackColor       =   &H00808080&
      Caption         =   "Economy Class"
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   4200
      Width           =   1935
   End
   Begin VB.OptionButton optBusinessClass 
      BackColor       =   &H00808080&
      Caption         =   "Business Class"
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   4200
      Width           =   1935
   End
   Begin VB.OptionButton optFirstClass 
      BackColor       =   &H00808080&
      Caption         =   "First Class"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label flight 
      BackColor       =   &H000000C0&
      Caption         =   "Choose your flight class."
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image imgplane 
      Height          =   3360
      Left            =   480
      Picture         =   "FrmFlight.frx":0000
      Top             =   600
      Width           =   5730
   End
   Begin VB.Image Image2 
      Height          =   15
      Left            =   1920
      Top             =   1800
      Width           =   255
   End
End
Attribute VB_Name = "FrmFlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'Project Name : Exotica_Travel (Amanda Whitcomb's VBProject.vbp)
'Form Name : frmFlight (frmFlight.frm)
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
Dim PATH As String



Private Sub cmdCompute_Click()
'declare variables
Dim flight(1 To 12) As Double

'Clear whatever may be in Results for repeated use.
results.Cls

'Determine which "Option" the user has selected.
If optFirstClass = True Then
    f = 1
ElseIf optBusinessClass = True Then
    f = 2
ElseIf optEconomyClass = True Then
    f = 3
End If

'Open the data file "file.txt" for the Arrays that
'are used in TravDest.
Open PATH & "flight.txt" For Input As #1

For i = 1 To 12
    Input #1, flight(i) 'Read data into the respective Arrays.
Next i
Close #1 'close the file

'Print the cost of the option the user has selected
'read from an array
If d = 1 And f = 1 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(1))
    ff = flight(1)
ElseIf d = 1 And f = 2 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(2))
    ff = flight(2)
ElseIf d = 1 And f = 3 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(3))
    ff = flight(3)
ElseIf d = 2 And f = 1 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(4))
    ff = flight(4)
ElseIf d = 2 And f = 2 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(5))
    ff = flight(5)
ElseIf d = 2 And f = 3 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(6))
    ff = flight(6)
ElseIf d = 3 And f = 1 Then
   results.Print "Your flight will cost "; FormatCurrency(flight(7))
   ff = flight(7)
ElseIf d = 3 And f = 2 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(8))
    ff = flight(8)
ElseIf d = 3 And f = 3 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(9))
    ff = flight(9)
ElseIf d = 4 And f = 1 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(10))
    ff = flight(10)
ElseIf d = 4 And f = 2 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(11))
    ff = flight(11)
ElseIf d = 4 And f = 3 Then
    results.Print "Your flight will cost "; FormatCurrency(flight(12))
    ff = flight(12)
End If

'Disable the "Compute" button
'and enable the "Continue" button.
cmdCompute.Enabled = False
cmdContinue.Enabled = True
End Sub

Private Sub cmdContinue_Click()
'Hide the Flight selection screen and show
'the Hotel selection screen for the users next input.
FrmFlight.Hide
FrmHotel.Show

'Disable the Continue button and Re-Enable
'the Compute button for repeated use.
cmdCompute.Enabled = True
cmdContinue.Enabled = False
End Sub

Private Sub Form_Load()
PATH = "m:\Amanda Whitcomb\Destination\"
End Sub

Private Sub optBusinessClass_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optEconomyClass_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optFirstClass_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub
