VERSION 5.00
Begin VB.Form FormFVAD 
   BackColor       =   &H0000FFFF&
   Caption         =   "Future Value of an Annuity Due"
   ClientHeight    =   3480
   ClientLeft      =   4050
   ClientTop       =   3180
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4950
   Begin VB.CommandButton cmdcalcFVAD 
      BackColor       =   &H00C000C0&
      Caption         =   "Calculate"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdbackFVAD 
      BackColor       =   &H00C000C0&
      Caption         =   "Go Back to Main Page"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtperiodsFVAD 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtrateFVAD 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton quitFVAD 
      BackColor       =   &H00C000C0&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtprincipleFVAD 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picResultsFVAD 
      BackColor       =   &H00FFC0FF&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
   End
   Begin VB.CommandButton cmdclearFVAD 
      BackColor       =   &H00C000C0&
      Caption         =   "Clear"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Please enter the number of annuity payments"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   $"FormAD.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter the amount of each annuity payment"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FormFVAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'FormFVAD(FormAD.frm)
'Author- Andrew Schou
'3/12/04
'The purpose of this form is to allow the user to enter the data that pertains to
'his or her investment.  The user will then find out the value of his or her inverstment.


'dimentioins of the variables
Dim periodFVAD As Single
Dim valueFVAD As String
Dim principleFVAD As Single
Dim periodsFVAD As Single
Dim rateFVAD As Single
'go back to the main page
Private Sub cmdbackFVAD_Click()
FormFVAD.Hide
Introduction.Show
End Sub
'what the variables are equal to and the value of the formula
Private Sub cmdcalcFVAD_Click()
If txtprincipleFVAD.DataChanged And txtrateFVAD.DataChanged And txtperiodsFVAD.DataChanged Then
    rateFVAD = (txtrateFVAD.Text * 0.01)
    principleFVAD = txtprincipleFVAD.Text
    periodsFVAD = txtperiodsFVAD.Text
    valueFVAD = principleFVAD * ((((1 + rateFVAD) ^ periodsFVAD) - 1) / rateFVAD) * (1 + rateFVAD)
    picResultsFVAD.Print "The future value of the annuity due is "; FormatCurrency(valueFVAD); "."
End If

End Sub
'clear the results in the picture window
Private Sub cmdclearFVAD_Click()
picResultsFVAD.Cls
End Sub



Private Sub quitFVAD_Click()
End
End Sub

