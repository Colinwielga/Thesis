VERSION 5.00
Begin VB.Form FrmTravelDestination 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox results 
      BackColor       =   &H00FFFF80&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   6795
      TabIndex        =   7
      Top             =   3960
      Width           =   6855
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Select"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H8000000A&
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   855
      Left            =   7320
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.OptionButton optAustralia 
      BackColor       =   &H00FFFF00&
      Caption         =   "Sydney, Australia"
      Height          =   615
      Left            =   6840
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton optBahamas 
      BackColor       =   &H00FFFF00&
      Caption         =   "Nassau, Bahamas"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton optJamaica 
      BackColor       =   &H00FFFF00&
      Caption         =   "Montego Bay, Jamaica"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton optMexico 
      BackColor       =   &H00FFFF00&
      Caption         =   "Acapulco, Mexico"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Image imgJamaica 
      Height          =   1980
      Left            =   2520
      Picture         =   "VB_Vacation.frx":0000
      Top             =   960
      Width           =   2010
   End
   Begin VB.Image imgBahamas 
      Height          =   1920
      Left            =   4680
      Picture         =   "VB_Vacation.frx":D092
      Top             =   960
      Width           =   2040
   End
   Begin VB.Image imgAustralia 
      Height          =   1920
      Left            =   6840
      Picture         =   "VB_Vacation.frx":19CD4
      Top             =   960
      Width           =   2115
   End
   Begin VB.Image imgMexico 
      Height          =   1935
      Left            =   120
      Picture         =   "VB_Vacation.frx":27116
      Top             =   960
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Choose Your Vacation Destination!"
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "FrmTravelDestination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Exotica_Travel (Amanda Whitcomb's VBProject.vbp)
'Form Name : frmTravelDestination (frmTravelDestination.frm)
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

'declare variables
Dim destination(1 To 4) As String
Dim temperature(1 To 4) As Double


Private Sub cmdCompute_Click()
'Clear whatever may be in Results for repeated use.
results.Cls

'Determine which "Option" the user has selected.
If optMexico = True Then
    d = 1
ElseIf optJamaica = True Then
    d = 2
ElseIf optBahamas = True Then
    d = 3
ElseIf optAustralia = True Then
    d = 4
End If

'Open the data file "destination.txt" for the Arrays that
'are used in TravDest.
Open PATH & "destination.txt" For Input As #1

For i = 1 To 4
    Input #1, destination(i), temperature(i) 'Read data into the respective Arrays.
Next i
Close #1 'close the file

'Print the option the user has selected
'along with the temperature of the destination
'read from an array
If d = 1 Then
    results.Print destination(1); "  The average temperature is"; temperature(1); "degrees F"
ElseIf d = 2 Then
    results.Print destination(2); "  The average temperature is "; temperature(2); "degrees F"
ElseIf d = 3 Then
    results.Print destination(3); "  The average temperature is"; temperature(3); "degrees F"
ElseIf d = 4 Then
    results.Print destination(4); "  The average temperature is"; temperature(4); "degrees F"
End If

'Disable the Compute button
'and enable the Continue button.
cmdCompute.Enabled = False
cmdContinue.Enabled = True
End Sub

Private Sub cmdContinue_Click()
Close #1
'Hide the Travel Destination selection screen and show
'the Flight selection screen for the users next input.
FrmTravelDestination.Hide
FrmFlight.Show

'Disable the Continue button and Re-Enable
'the Compute button for repeated use.
cmdCompute.Enabled = True
cmdContinue.Enabled = False
End Sub

Private Sub Form_Load()
PATH = "m:\Amanda Whitcomb\Destination\"
End Sub

Private Sub optAustralia_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optBahamas_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optJamaica_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub

Private Sub optMexico_Click()
'Enable Compute button after a selection has been made.
cmdCompute.Enabled = True
End Sub
