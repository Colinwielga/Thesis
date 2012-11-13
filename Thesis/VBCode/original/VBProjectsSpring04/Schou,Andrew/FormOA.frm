VERSION 5.00
Begin VB.Form FormFVOA 
   BackColor       =   &H000080FF&
   Caption         =   "Future Value of an Ordinary Annuity"
   ClientHeight    =   3450
   ClientLeft      =   4050
   ClientTop       =   3180
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4920
   Begin VB.CommandButton cmdclearFVOA 
      BackColor       =   &H00FF8080&
      Caption         =   "Clear"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox picResultsFVOA 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   2160
      Width           =   4695
   End
   Begin VB.TextBox txtprincipleFVOA 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton quitFVOA 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtrateFVOA 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtperiodsFVOA 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdbackFVOA 
      BackColor       =   &H00FF8080&
      Caption         =   "Go Back to Main Page"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdcalcFVOA 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "Enter the amount of each annuity payment"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   $"FormOA.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "Please enter the number of annuity payments"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "FormFVOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'FormFVOA(FormOA.frm)
'Author- Andrew Schou
'3/12/04
'The purpose of this form is to allow the user to enter the data that pertains to
'his or her investment.  The user will then find out the value of his or her inverstment.

'dimentioins of the variables
Dim periodFVOA As Single
Dim valueFVOA As String
Dim principleFVOA As Single
Dim periodsFVOA As Single
Dim rateFVOA As Single
'go back to the main page
Private Sub cmdbackFVOA_Click()
FormFVOA.Hide
Introduction.Show
End Sub
'what the variables are equal to and the value of the formula
Private Sub cmdcalcFVOA_Click()
If txtprincipleFVOA.DataChanged And txtrateFVOA.DataChanged And txtperiodsFVOA.DataChanged Then
    rateFVOA = (txtrateFVOA.Text * 0.01)
    principleFVOA = txtprincipleFVOA.Text
    periodsFVOA = txtperiodsFVOA.Text
    valueFVOA = principleFVOA * (((1 + rateFVOA) ^ periodsFVOA) - 1) / rateFVOA
    picResultsFVOA.Print "The future value of the annuity  is "; FormatCurrency(valueFVOA); "."
End If


End Sub
'clear the results in the picture window
Private Sub cmdclearFVOA_Click()
picResultsFVOA.Cls
End Sub



Private Sub quitFVOA_Click()
End
End Sub

