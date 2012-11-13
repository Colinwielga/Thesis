VERSION 5.00
Begin VB.Form FormPVAD 
   BackColor       =   &H0080FFFF&
   Caption         =   "Present Value of an Annuity Due"
   ClientHeight    =   3435
   ClientLeft      =   4980
   ClientTop       =   2880
   ClientWidth     =   4905
   LinkTopic       =   "Form3"
   ScaleHeight     =   3435
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdcalcPVAD 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdbackPVAD 
      BackColor       =   &H00FF8080&
      Caption         =   "Go Back to Main Page"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtperiodsPVAD 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtratePVAD 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton quitPVAD 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtprinciplePVAD 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picResultsPVAD 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
   End
   Begin VB.CommandButton cmdclearPVAD 
      BackColor       =   &H00FF8080&
      Caption         =   "Clear"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Please enter the number of annuity payments."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   $"FormPVAD.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Enter the amount of each annuity payment"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FormPVAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'FormPVAD(FormPVAD.frm)
'Author- Andrew Schou
'3/12/04
'The purpose of this form is to allow the user to enter the data that pertains to
'his or her investment.  The user will then find out the value of his or her inverstment.

'dimentioins of the variables
Dim periodPVAD As Single
Dim valuePVAD As String
Dim principlePVAD As Single
Dim periodsPVAD As Single
Dim ratePVAD As Single

'go back to the main page
Private Sub cmdbackPVAD_Click()
FormPVAD.Hide
Introduction.Show
End Sub
'what the variables are equal to and the value of the formula
Private Sub cmdcalcPVAD_Click()
If txtprinciplePVAD.DataChanged And txtratePVAD.DataChanged And txtperiodsPVAD.DataChanged Then
    ratePVAD = (txtratePVAD.Text * 0.01)
    principlePVAD = txtprinciplePVAD.Text
    periodsPVAD = txtperiodsPVAD.Text
    valuePVAD = principlePVAD * (1 + ((1 - ((1 + ratePVAD) ^ (-(periodsPVAD - 1)))) / ratePVAD))
    picResultsPVAD.Print "The present value of the annuity due is "; FormatCurrency(valuePVAD); "."
End If

End Sub
'clear the results in the picture window
Private Sub cmdclearPVAD_Click()
picResultsPVAD.Cls
End Sub



Private Sub quitPVAD_Click()
End
End Sub
