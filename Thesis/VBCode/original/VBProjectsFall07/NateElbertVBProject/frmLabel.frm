VERSION 5.00
Begin VB.Form frmLabel 
   Caption         =   "Label for Student Tax Return"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9960
      TabIndex        =   1
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   7185
      Left            =   3360
      Picture         =   "frmLabel.frx":0000
      Top             =   720
      Width           =   6480
   End
   Begin VB.Label lblHeader 
      Caption         =   "Estimated Income Tax Return for Single Filers With No Dependents "
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   15855
   End
End
Attribute VB_Name = "frmLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBegin_Click()
UserName = InputBox("Please enter your full name with middle initial.")             'User is asked to input personal information
Address = InputBox("Please enter your home address(number and street).")
City = InputBox("Please enter your home city.")
State = InputBox("Please enter your home state.")
ZIPcode = InputBox("Please enter your ZIP code.")
    frmIncome.Show                                                                  'Button hides the Label form and shows the Income form
    frmLabel.Hide
End Sub

Private Sub cmdContinue_Click()
    frmIncome.Show                                                                  'Button hides the Income form and shows the Label form
    frmLabel.Hide
End Sub

Private Sub cmdQuit_Click()
End                                                                                 'Button ends the program
End Sub

