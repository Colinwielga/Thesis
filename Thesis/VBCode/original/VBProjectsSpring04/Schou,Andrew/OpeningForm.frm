VERSION 5.00
Begin VB.Form OpeningForm 
   BackColor       =   &H0080C0FF&
   Caption         =   "Your Own Investment"
   ClientHeight    =   3720
   ClientLeft      =   4050
   ClientTop       =   3180
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6855
   Begin VB.CommandButton cmdgoback 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdAD 
      BackColor       =   &H008080FF&
      Caption         =   "Aunnity Due"
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOA 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ordinary Annuity"
      Height          =   615
      Left            =   1440
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton quitMain 
      BackColor       =   &H0080FF80&
      Caption         =   "Quit"
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSP 
      BackColor       =   &H00FF80FF&
      Caption         =   "Single Payment"
      Height          =   615
      Left            =   120
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label statement 
      BackColor       =   &H0080C0FF&
      Caption         =   "Please Enter the type of payment you plan to use."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   6495
   End
   Begin VB.Label Program 
      BackColor       =   &H0080C0FF&
      Caption         =   $"OpeningForm.frx":0000
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "OpeningForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'OpeningForm(Opening.frm)
'Author- Andrew Schou
'3/11/04
'The purpose of this form is to introduce the user to this portion of the program.  It also gives the user the option of
'picking which type of investment they are going to figure out.


'each button will clear the current window and take the  user to the window of the type of investment thar he/she selected
Private Sub cmdAD_Click()
FormAD.Show
OpeningForm.Hide
End Sub

Private Sub cmdgoback_Click()
OpeningForm.Hide
Introduction.Show
End Sub

Private Sub cmdOA_Click()
FormOA.Show
OpeningForm.Hide
End Sub

Private Sub cmdSP_Click()
FormSS.Show
OpeningForm.Hide
End Sub


Private Sub quitMain_Click()
End
End Sub
