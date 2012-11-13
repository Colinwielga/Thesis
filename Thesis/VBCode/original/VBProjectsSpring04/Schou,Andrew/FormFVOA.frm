VERSION 5.00
Begin VB.Form FormAD 
   BackColor       =   &H008080FF&
   Caption         =   "Annuity Due"
   ClientHeight    =   1710
   ClientLeft      =   4665
   ClientTop       =   3780
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   1710
   ScaleWidth      =   4440
   Begin VB.CommandButton cancelAD 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton presentAD 
      BackColor       =   &H0080FFFF&
      Caption         =   "Present Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton quitAD 
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton futureAD 
      BackColor       =   &H0080FFFF&
      Caption         =   "Future Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
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
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "FormAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'FormAD(FormFVOA.frm)
'Author- Andrew Schou
'3/11/04
'The purpose of this form is to give the user the option to pick whether he/she wants to
'find the present of future value of the investment.

'each button will hide the current window, and take the user to the appropriate window for the type of investment they selected
Private Sub cancelAD_Click(Index As Integer)
FormAD.Hide
Introduction.Show
End Sub

Private Sub futureAD_Click(Index As Integer)
FormAD.Hide
FormFVAD.Show
End Sub

Private Sub presentAD_Click(Index As Integer)
FormAD.Hide
FormPVAD.Show
End Sub

Private Sub quitAD_Click(Index As Integer)
End
End Sub
