VERSION 5.00
Begin VB.Form ApplicationForm 
   BackColor       =   &H00FFFF00&
   Caption         =   "Loan Application"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00FFFF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FF8080&
      Caption         =   "Clear"
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   1455
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   2655
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3720
      ScaleHeight     =   1395
      ScaleWidth      =   7995
      TabIndex        =   21
      Top             =   4920
      Width           =   8055
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00FF8080&
      Caption         =   "Submit"
      Height          =   1455
      Left            =   360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox txtAmount 
      Height          =   615
      Left            =   3720
      TabIndex        =   19
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txtZip 
      Height          =   495
      Left            =   9840
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtState 
      Height          =   495
      Left            =   8880
      TabIndex        =   12
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtCity 
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtStreet 
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtPhone 
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txtLast 
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtFirst 
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtMr 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblDesign 
      BackColor       =   &H00FFFF00&
      Caption         =   "Designed By Brian Finley"
      Height          =   255
      Left            =   9240
      TabIndex        =   24
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Please Enter The Amount Of Money You Desire To Be Loaned"
      Height          =   615
      Left            =   480
      TabIndex        =   18
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label lblZip 
      BackColor       =   &H00FFFF00&
      Caption         =   "Zip Code"
      Height          =   255
      Left            =   9960
      TabIndex        =   17
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblState 
      BackColor       =   &H00FFFF00&
      Caption         =   "State"
      Height          =   255
      Left            =   8880
      TabIndex        =   16
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblCity 
      BackColor       =   &H00FFFF00&
      Caption         =   "City"
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblStreet 
      BackColor       =   &H00FFFF00&
      Caption         =   "Street/PO Box"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Please Enter Your Address"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lblLast 
      BackColor       =   &H00FFFF00&
      Caption         =   "Last"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblFirst 
      BackColor       =   &H00FFFF00&
      Caption         =   "First"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblMr 
      BackColor       =   &H00FFFF00&
      Caption         =   "Mr./Mrs."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00FFFF00&
      Caption         =   "Please Enter Your Phone Number"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFF00&
      Caption         =   "Please Enter Your Name"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "ApplicationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name is LoanVBProject(FinalVBProject.vbp)
'The form is named ApplicationForm(VBProject.frm)
'This form was written by Brian Finley on 3/14/04
'This form is an application for a loan
'It gets personal information from the user and then confirms that the information was submitted

Option Explicit
Dim FirstName As String, LastName As String, MrMrs As String
Dim Phone As String, Address As String, City As String, State As String
Dim Zip As Double, Amount As Double


'Clears the picture box
Private Sub cmdClear_Click()
picResults2.Cls
End Sub

'Hides ApplicationForm and Shows LoanForm
Private Sub cmdQuit_Click()
LoanForm.Show
ApplicationForm.Hide
End Sub

'Prints out a letter the the user verifying the information
'The If-Then makes sure that all of the blanks are filled out before the submit button will print anything out
Private Sub cmdSubmit_Click()
If txtFirst <> "" And txtLast <> "" And txtMr <> "" And txtPhone <> "" And txtStreet <> "" And txtCity <> "" And txtState <> "" And txtAmount <> "" And txtZip <> "" Then
    picResults2.Cls
    FirstName = txtFirst
    LastName = txtLast
    MrMrs = txtMr
    Phone = txtPhone
    Address = txtStreet
    City = txtCity
    State = txtState
    Zip = txtZip
    If txtAmount >= 1000 Then       'Makes sure that the loan is for the minimum of $1000
        Amount = txtAmount
        picResults2.Print "Thank you"; " "; MrMrs; ". "; FirstName; " "; LastName; " "; "for applying for a loan of"; " "; FormatCurrency(Amount); "."
        picResults2.Print "You will be contacted by phone later today at"; " "; Phone; " "; "to set up an appointment with a loan officer."
        picResults2.Print "You will also be mailed more information about our loans at"; " "; Address; " "; City; ","; " "; State; " "; Zip; "."
        picResults2.Print "Have a great day!"
    Else        'gives the user an error message if the loan is not big enough
        MsgBox "Your loan must be for at least $1,000", , "Error"
    End If
Else            'If one of the blanks is not filled out it triggers an error message
    MsgBox "Please fill in all of the blanks.", , "Error"
End If
End Sub


