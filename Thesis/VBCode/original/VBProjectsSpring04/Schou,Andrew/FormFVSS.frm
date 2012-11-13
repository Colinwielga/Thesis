VERSION 5.00
Begin VB.Form FormOA 
   BackColor       =   &H00FFFF80&
   Caption         =   "Ordinary Annuity"
   ClientHeight    =   1665
   ClientLeft      =   5595
   ClientTop       =   4695
   ClientWidth     =   4425
   LinkTopic       =   "Form3"
   ScaleHeight     =   1665
   ScaleWidth      =   4425
   Begin VB.CommandButton futureOA 
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
   Begin VB.CommandButton quitOA 
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
   Begin VB.CommandButton presentOA 
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
   Begin VB.CommandButton cancelOA 
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
      BackColor       =   &H00FFFF80&
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
Attribute VB_Name = "FormOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelAD_Click(Index As Integer)
FormAD.Hide
OpeningForm.Show
End Sub
'Investments(project1.vbp)
'FormOA(FormFVSS.frm)
'Author- Andrew Schou
'3/11/04
'The purpose of this form is to give the user the option to pick whether he/she wants to
'find the present of future value of the investment.


'each button will hide the current window, and take the user to the appropriate window for the type of investment they selected
Private Sub futureOA_Click(Index As Integer)
FormOA.Hide
FormFVOA.Show
End Sub

Private Sub presentOA_Click(Index As Integer)
FormOA.Hide
FormPVOA.Show
End Sub

Private Sub quitOA_Click(Index As Integer)
End
End Sub

Private Sub cancelOA_Click(Index As Integer)
FormOA.Hide
Introduction.Show
End Sub
