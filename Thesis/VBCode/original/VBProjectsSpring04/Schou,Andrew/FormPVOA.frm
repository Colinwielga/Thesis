VERSION 5.00
Begin VB.Form FormPVOA 
   BackColor       =   &H00FF80FF&
   Caption         =   "Present Value of Ordinary Annuity"
   ClientHeight    =   3435
   ClientLeft      =   4980
   ClientTop       =   3180
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   ScaleHeight     =   3435
   ScaleWidth      =   4890
   Begin VB.CommandButton cmdcalcPVOA 
      BackColor       =   &H000080FF&
      Caption         =   "Calculate"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdbackPVOA 
      BackColor       =   &H000080FF&
      Caption         =   "Go Back to Main Page"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtperiodsPVOA 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtratePVOA 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton quitPVOA 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtprinciplePVOA 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picResultsPVOA 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
   End
   Begin VB.CommandButton cmdclearPVOA 
      BackColor       =   &H000080FF&
      Caption         =   "Clear"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Please enter the number of annuity payments."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF80FF&
      Caption         =   $"FormPVOA.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF80FF&
      Caption         =   "Enter the amount of each annuity payment"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FormPVOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'FormPVOA(FormPVOA.frm)
'Author- Andrew Schou
'3/12/04
'The purpose of this form is to allow the user to enter the data that pertains to
'his or her investment.  The user will then find out the value of his or her inverstment.

'dimentioins of the variables
Dim periodPVOA As Single
Dim valuePVOA As String
Dim principlePVOA As Single
Dim periodsPVOA As Single
Dim ratePVOA As Single
'go back to the main page
Private Sub cmdbackPVOA_Click()
FormPVOA.Hide
Introduction.Show
End Sub
'what the variables are equal to and the value of the formula
Private Sub cmdcalcPVOA_Click()
If txtprinciplePVOA.DataChanged And txtratePVOA.DataChanged And txtperiodsPVOA.DataChanged Then
    ratePVOA = (txtratePVOA.Text * 0.01)
    principlePVOA = txtprinciplePVOA.Text
    periodsPVOA = txtperiodsPVOA.Text
    valuePVOA = principlePVOA * ((1 - ((1 + ratePVOA) ^ (-periodsPVOA))) / ratePVOA)
    picResultsPVOA.Print "The present value of the ordinary annuity is "; FormatCurrency(valuePVOA); "."
End If

End Sub
'clear the results in the picture window
Private Sub cmdclearPVOA_Click()
picResultsPVOA.Cls
End Sub



Private Sub quitPVOA_Click()
End
End Sub

