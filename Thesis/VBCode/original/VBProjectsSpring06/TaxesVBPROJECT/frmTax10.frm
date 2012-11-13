VERSION 5.00
Begin VB.Form frmTax10 
   BackColor       =   &H80000013&
   Caption         =   "Line 10 - Find your Tax"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForm 
      Caption         =   "Return to Tax Form"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
   End
   Begin VB.PictureBox pictax 
      Height          =   2055
      Left            =   1080
      ScaleHeight     =   1995
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter your income"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label lblMyName 
      Caption         =   "By Brent Mergen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label lblTax 
      BackColor       =   &H80000013&
      Caption         =   "TAX"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmTax10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'Line 10 - Find Your tax (frmTax10)
'Brent Timothy Mergen
'24 March 2006
'Calculates Question 10 from the Tax Return form.

Option Explicit
Dim income(1 To 500) As Single
Dim tax(1 To 500) As Single
Dim pos As Integer
Dim entryincome As Single
Dim number As Integer

Private Sub cmdForm_Click()
    frmTax10.Hide
    frmTaxInput.Show
End Sub


Private Sub Command1_Click()
    Open App.Path & "\tax.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, income(pos), tax(pos)
    Loop
    Close #1
    pictax.Cls
    pos = 1
    Do Until income(pos) > answer6
        number = pos
        pos = pos + 1
    Loop
    pictax.Print tax(pos + 1)
    Overalltax = tax(pos + 1)
End Sub
