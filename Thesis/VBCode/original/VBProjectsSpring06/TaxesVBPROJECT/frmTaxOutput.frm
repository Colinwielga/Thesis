VERSION 5.00
Begin VB.Form frmTaxOutput 
   BackColor       =   &H80000013&
   Caption         =   "Tax Return"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinished 
      Caption         =   "Finished With My Taxes!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "View my Tax return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.PictureBox picTaxOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1680
      ScaleHeight     =   1635
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H80000013&
      Caption         =   "By Brent Mergen"
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "frmTaxOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'Tax Return (frmTaxOutput)
'Brent Timothy Mergen
'24 March 2006
'Shows you your routing and account numbers, along with that it displays where your refund is going to.

Private Sub cmdFinished_Click()
    frmTaxOutput.Hide
    frmFrontpage.Show
    MsgBox "Thank you for using Form 1040EZ with Brent Mergen and Associates.  You are finished!", , "Finished!"
End Sub

Private Sub cmdview_Click()
    picTaxOutput.Print "Your account number is "; accountnumber
    picTaxOutput.Print "Your routing Number is "; routingnumber
    picTaxOutput.Print "Your tax return will be electronically deposited to your "; selection; " account."
End Sub
