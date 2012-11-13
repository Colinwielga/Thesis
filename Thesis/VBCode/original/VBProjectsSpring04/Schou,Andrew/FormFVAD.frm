VERSION 5.00
Begin VB.Form FormSS 
   BackColor       =   &H00FF80FF&
   Caption         =   "Single Sum"
   ClientHeight    =   1635
   ClientLeft      =   5280
   ClientTop       =   4695
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   4425
   Begin VB.CommandButton futureSS 
      BackColor       =   &H00FF8080&
      Caption         =   "Future Value"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton quitSS 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      Height          =   495
      Index           =   1
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton presentSS 
      BackColor       =   &H00FF8080&
      Caption         =   "Present Value"
      Height          =   495
      Index           =   2
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cancelSS 
      BackColor       =   &H00FF8080&
      Caption         =   "Cancel"
      Height          =   495
      Index           =   3
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Please select whether you would like to find the future or present value of the investment."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "FormSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'FormSS(FormFVAD.frm)
'Author- Andrew Schou
'3/11/04
'The purpose of this form is to give the user the option to pick whether he/she wants to
'find the present of future value of the investment.

'each button will hide the current window, and take the user to the appropriate window for the type of investment they selected
Private Sub cancelSS_Click(Index As Integer)
FormSS.Hide
Introduction.Show
End Sub

Private Sub futureSS_Click(Index As Integer)
FormSS.Hide
FormFVSS.Show
End Sub

Private Sub presentSS_Click(Index As Integer)
FormSS.Hide
FormPVSS.Show
End Sub

Private Sub quitSS_Click(Index As Integer)
End
End Sub


