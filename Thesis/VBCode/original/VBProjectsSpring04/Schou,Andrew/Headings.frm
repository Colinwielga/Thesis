VERSION 5.00
Begin VB.Form Headings 
   Caption         =   "Explination of the Headings"
   ClientHeight    =   4335
   ClientLeft      =   7125
   ClientTop       =   6210
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   5910
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Close"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   $"Headings.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   $"Headings.frx":00EE
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "Date of Maturity- This is the date at which the bond will mature."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "Interest Rate- This is the interest rate at which the bond pays interest each year."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Company- This is the name of the company that issued the bonds listed."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Headings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'Headings(Headings.frm)
'Author- Andrew Schou
'3/14/04
'The purpose of this form is to tell the user what each caption meens in the
'tables on the previous form.

'this button closes the window
Private Sub cmdClose_Click()
Headings.Hide
End Sub
