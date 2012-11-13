VERSION 5.00
Begin VB.Form FormFVSS 
   BackColor       =   &H0080FF80&
   Caption         =   "Future Value of a Single Sum"
   ClientHeight    =   3480
   ClientLeft      =   4365
   ClientTop       =   3180
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4935
   Begin VB.CommandButton cmdclearFVSS 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox picResultsFVSS 
      BackColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   9
      Top             =   2160
      Width           =   4695
   End
   Begin VB.TextBox txtprincipleFVSS 
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton quitFVSS 
      BackColor       =   &H008080FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtrateFVSS 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtperiodsFVSS 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdbackFVSS 
      BackColor       =   &H008080FF&
      Caption         =   "Go Back to Main Page"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdcalcFVSS 
      BackColor       =   &H008080FF&
      Caption         =   "Calculate"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Enter the amount of the Single Sum Payment"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   $"FormSS.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Please enter the number of interest periods."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "FormFVSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'FormFVSS(FormSS.frm)
'Author- Andrew Schou
'3/12/04
'The purpose of this form is to allow the user to enter the data that pertains to
'his or her investment.  The user will then find out the value of his or her inverstment.

'dimentioins of the variables
Dim periodFVSS As Single
Dim valueFVSS As String
Dim principleFVSS As Single
Dim periodsFVSS As Single
Dim rateFVSS As Single
'go back to the main page
Private Sub cmdbackFVSS_Click()
FormFVSS.Hide
Introduction.Show
End Sub
'what the variables are equal to and the value of the formula
Private Sub cmdcalcFVSS_Click()
If txtprincipleFVSS.DataChanged And txtrateFVSS.DataChanged And txtperiodsFVSS.DataChanged Then
    rateFVSS = (txtrateFVSS.Text * 0.01)
    principleFVSS = txtprincipleFVSS.Text
    periodsFVSS = txtperiodsFVSS.Text
    valueFVSS = principleFVSS * ((1 + rateFVSS) ^ periodsFVSS)
    picResultsFVSS.Print "The future value of the investment is "; FormatCurrency(valueFVSS); "."
End If

End Sub
'clear the results in the picture window
Private Sub cmdclearFVSS_Click()
picResultsFVSS.Cls
End Sub



Private Sub quitFVSS_Click()
End
End Sub
