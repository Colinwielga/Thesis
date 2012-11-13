VERSION 5.00
Begin VB.Form FrmHotel 
   BackColor       =   &H0080FFFF&
   Caption         =   "Hotel"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox results 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   360
      ScaleHeight     =   675
      ScaleWidth      =   5715
      TabIndex        =   6
      Top             =   4560
      Width           =   5775
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Calculate the cost of your hotel."
      Enabled         =   0   'False
      Height          =   1095
      Left            =   6480
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   8640
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin VB.OptionButton optThreeStar 
      BackColor       =   &H0080FFFF&
      Caption         =   "Three Star Hotel"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
   End
   Begin VB.OptionButton optTwoStar 
      BackColor       =   &H0080FFFF&
      Caption         =   "Two Star Hotel"
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
   End
   Begin VB.OptionButton optFourStar 
      BackColor       =   &H0080FFFF&
      Caption         =   "Four Star Hotel"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
   End
   Begin VB.OptionButton optFiveStar 
      BackColor       =   &H0080FFFF&
      Caption         =   "Five Star Hotel"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Choose a Hotel"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image imgtwostar 
      Height          =   1950
      Left            =   8160
      Picture         =   "FrmHotel.frx":0000
      Top             =   1320
      Width           =   2595
   End
   Begin VB.Image imgthreestar 
      Height          =   2475
      Left            =   5880
      Picture         =   "FrmHotel.frx":10852
      Top             =   840
      Width           =   1860
   End
   Begin VB.Image imgfourstar 
      Height          =   1725
      Left            =   2880
      Picture         =   "FrmHotel.frx":1F858
      Top             =   1560
      Width           =   2580
   End
   Begin VB.Image imgfivestar 
      Height          =   2070
      Left            =   240
      Picture         =   "FrmHotel.frx":2E066
      Top             =   1200
      Width           =   2325
   End
End
Attribute VB_Name = "FrmHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declare variables
Dim hotel(1 To 4) As Double

'Project Name : Exotica_Travel (Amanda Whitcomb's VBProject.vbp)
'Form Name : frmHotel (frmHotel.frm)
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

Private Sub cmdCompute_Click()
'Clear whatever may be in Results for repeated use.
results.Cls

'Determine which "Option" the user has selected.
If optFiveStar = True Then
    h = 1
ElseIf optFourStar = True Then
    h = 2
ElseIf optThreeStar = True Then
    h = 3
ElseIf optTwoStar = True Then
    h = 4
End If

'Open the data file "hotel.txt" for the Arrays that
'are used in TravDest.
Open path & "hotel.txt" For Input As #1

For i = 1 To 4
    Input #1, hotel(i) 'Read data into the respective Arrays.
Next i
Close #1 'close the file

'Print the cost of the option the user has selected
'read from an array
If h = 1 Then
    results.Print "Your hotel stay costs "; FormatCurrency(hotel(1))
    hh = hotel(1)
ElseIf h = 2 Then
    results.Print "Your hotel stay costs "; FormatCurrency(hotel(2))
    hh = hotel(2)
ElseIf h = 3 Then
    results.Print "Your hotel stay costs "; FormatCurrency(hotel(3))
    hh = hotel(3)
ElseIf h = 4 Then
    results.Print "Your hotel stay costs "; FormatCurrency(hotel(4))
    hh = hotel(4)
End If

'Disable the Compute button
'and enable the Continue button.
cmdCompute.Enabled = False
cmdContinue.Enabled = True

End Sub

Private Sub cmdContinue_Click()
Close #1
'Hide the Hotel selection screen and show
'the Entertainment selection screen for the users next input.
FrmHotel.Hide
FrmEntertainment.Show

'Disable the "Next Selection" button and Re-Enable
'the "Print Selection" button for repeated use.
cmdCompute.Enabled = True
cmdContinue.Enabled = False
End Sub

Private Sub optFiveStar_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optFourStar_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optThreeStar_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optTwoStar_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub
