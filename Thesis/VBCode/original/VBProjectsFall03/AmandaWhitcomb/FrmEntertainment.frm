VERSION 5.00
Begin VB.Form FrmEntertainment 
   BackColor       =   &H00FF8080&
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute your entertainment cost"
      Height          =   1095
      Left            =   6000
      TabIndex        =   7
      Top             =   4440
      Width           =   1815
   End
   Begin VB.PictureBox results 
      BackColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   4320
      Width           =   5415
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   7920
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.OptionButton optClub 
      BackColor       =   &H00FF8080&
      Caption         =   "Go to a Nightclub"
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OptionButton optSnorkel 
      BackColor       =   &H00FF8080&
      Caption         =   "Snorkel"
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.OptionButton optRelax 
      BackColor       =   &H00FF8080&
      Caption         =   "Relax"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.OptionButton optSwim 
      BackColor       =   &H00FF8080&
      Caption         =   "Swim with Dolphins"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image imgclub 
      Height          =   1830
      Left            =   6960
      Picture         =   "FrmEntertainment.frx":0000
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Image imgsnorkel 
      Height          =   1470
      Left            =   4320
      Picture         =   "FrmEntertainment.frx":85B2
      Top             =   1920
      Width           =   2205
   End
   Begin VB.Image imgrelax 
      Height          =   1455
      Left            =   1920
      Picture         =   "FrmEntertainment.frx":12FEC
      Top             =   1920
      Width           =   2145
   End
   Begin VB.Image imgdolphin 
      Height          =   1725
      Left            =   240
      Picture         =   "FrmEntertainment.frx":1D3DE
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Label Entertainment 
      BackColor       =   &H00FF8080&
      Caption         =   "Choose your entertainment."
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "FrmEntertainment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Travel Destination (Amanda Whitcomb's VB Project.vbp)
'Form Name : frmEntertainment (Entertainment.frm)
'Author: Amanda Whitcomb
'Date Written: October 30, 2003
'Purpose of Form:
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

'Declare variables
Dim enter(1 To 4) As Double


'Project Name : Exotica_Travel (Amanda Whitcomb's VBProject.vbp)
'Form Name : frmEntertainment (frmEntertainment.frm)
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
If optSwim = True Then
    e = 1
ElseIf optRelax = True Then
    e = 2
ElseIf optSnorkel = True Then
    e = 3
ElseIf optClub = True Then
    e = 4
End If

'Open the data file "entertainment.txt" for the Arrays that
'are used in TravDest.
Open path & "entertainment.txt" For Input As #1

For i = 1 To 4
    Input #1, enter(i) 'read data into the respective arrays
Next i
Close #1 'close the file

'Print the cost of the option the user has selected
'read from an array
If e = 1 Then
    results.Print "Your entertainment costs "; FormatCurrency(enter(1))
    ee = enter(1)
ElseIf e = 2 Then
    results.Print "Your entertainment costs "; FormatCurrency(enter(2))
    ee = enter(2)
ElseIf e = 3 Then
    results.Print "Your entertainment costs "; FormatCurrency(enter(3))
    ee = enter(3)
ElseIf e = 4 Then
    results.Print "Your entertainment costs "; FormatCurrency(enter(4))
    ee = enter(4)
End If

'Disable the compute button
'and enable the continue button
cmdCompute.Enabled = False
cmdContinue.Enabled = True
End Sub

Private Sub cmdContinue_Click()
Close #1
'Hide the Entertainment selection screen and show
'the Meal selection screen for the users next input.
FrmEntertainment.Hide
frmMeal.Show

'Disable the Continue button and Re-Enable
'the Compute button for repeated use.
cmdCompute.Enabled = True
cmdContinue.Enabled = False
End Sub

Private Sub Option1_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optClub_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optRelax_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optSnorkel_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optSwim_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub
