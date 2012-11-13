VERSION 5.00
Begin VB.Form FormPVSS 
   BackColor       =   &H00FF8080&
   Caption         =   "Persent Value of a Single Sum"
   ClientHeight    =   3435
   ClientLeft      =   4980
   ClientTop       =   3180
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4875
   Begin VB.CommandButton cmdcalcPVSS 
      BackColor       =   &H0080FF80&
      Caption         =   "Calculate"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdbackPVSS 
      BackColor       =   &H0080FF80&
      Caption         =   "Go Back to Main Page"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtperiodsPVSS 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtratePVSS 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton quitPVSS 
      BackColor       =   &H0080FF80&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtprinciplePVSS 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picResultsPVSS 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
   End
   Begin VB.CommandButton cmdclearPVSS 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Please enter the number of interest periods."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   $"FormPVSS.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Enter the amount of the Single Sum Payment"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FormPVSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'FormPVSS(FormPVSS.frm)
'Author- Andrew Schou
'3/12/04
'The purpose of this form is to allow the user to enter the data that pertains to
'his or her investment.  The user will then find out the value of his or her inverstment.

'dimentioins of the variables
Dim periodPVSS As Single
Dim valuePVSS As String
Dim principlePVSS As Single
Dim periodsPVSS As Single
Dim ratePVSS As Single
'go back to the main page
Private Sub cmdbackPVSS_Click()
FormPVSS.Hide
Introduction.Show
End Sub
'what the variables are equal to and the value of the formula
Private Sub cmdcalcPVSS_Click()
If txtprinciplePVSS.DataChanged And txtratePVSS.DataChanged And txtperiodsPVSS.DataChanged Then
    ratePVSS = (txtratePVSS.Text * 0.01)
    principlePVSS = txtprinciplePVSS.Text
    periodsPVSS = txtperiodsPVSS.Text
    valuePVSS = principlePVSS * ((1 + ratePVSS) ^ (-periodsPVSS))
    picResultsPVSS.Print "The present value of the investment is "; FormatCurrency(valuePVSS); "."
End If

End Sub
'clear the results in the picture window
Private Sub cmdclearPVSS_Click()
picResultsPVSS.Cls
End Sub



Private Sub quitPVSS_Click()
End
End Sub

