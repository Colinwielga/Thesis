VERSION 5.00
Begin VB.Form frmProfilePedo2 
   BackColor       =   &H00000000&
   Caption         =   "Pedophile Types"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdgo 
      BackColor       =   &H0000FF00&
      Caption         =   "Generate Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8760
      Width           =   2895
   End
   Begin VB.CommandButton cmdfix 
      BackColor       =   &H0000FFFF&
      Caption         =   "Fixated Molester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdMy 
      BackColor       =   &H0000FFFF&
      Caption         =   "Mysoped Molester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdInad 
      BackColor       =   &H0000FFFF&
      Caption         =   "Inadequate Molester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSI 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sexually Indiscriminate Molester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdMI 
      BackColor       =   &H0000FFFF&
      Caption         =   "Morally Indiscriminate Molester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdReg 
      BackColor       =   &H0000FFFF&
      Caption         =   "Regressed Molester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdreview 
      BackColor       =   &H00FF00FF&
      Caption         =   "Review Case"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please click on the button which corresponds to the typology you believe the offender to be."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   10935
   End
End
Attribute VB_Name = "frmProfilePedo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form uses command buttons so the user can select what they
'believe the offenders real typology to be.


Private Sub cmdfix_Click()
    pedo2answer = "six" 'this allows the user to click on a comand button to choose their profile
    frmProfilePedo2.Hide 'hides the current form
    frmPedo2Solve.Show 'takes user to final form
End Sub

Private Sub cmdInad_Click()
    pedo2answer = "four" 'this allows the user to click on a comand button to choose their profile
    frmProfilePedo2.Hide 'hides the current form
    frmPedo2Solve.Show 'takes user to final form
End Sub

Private Sub cmdMI_Click()
    pedo2answer = "two" 'this allows the user to click on a comand button to choose their profile
    frmProfilePedo2.Hide 'hides the current form
    frmPedo2Solve.Show 'takes user to final form
End Sub

Private Sub cmdMy_Click()
    pedo2answer = "five" 'this allows the user to click on a comand button to choose their profile
    frmProfilePedo2.Hide 'hides the current form
    frmPedo2Solve.Show 'takes user to final form
End Sub

Private Sub cmdReg_Click()
    pedo2answer = "one" 'this allows the user to click on a comand button to choose their profile
    frmProfilePedo2.Hide 'hides the current form
    frmPedo2Solve.Show 'takes user to final form
End Sub

Private Sub cmdreview_Click()
'Allows the user one last chance to review the case
    frmreviewcase4b.Show
End Sub

Private Sub cmdSI_Click()
    pedo2answer = "three" 'this allows the user to click on a comand button to choose their profile
    frmProfilePedo2.Hide 'hides the current form
    frmPedo2Solve.Show 'takes user to final form
End Sub


