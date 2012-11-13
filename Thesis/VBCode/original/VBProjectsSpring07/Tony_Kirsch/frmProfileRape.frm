VERSION 5.00
Begin VB.Form frmProfileRape 
   BackColor       =   &H00000000&
   Caption         =   "Rapist Typologies"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAE 
      BackColor       =   &H0000FFFF&
      Caption         =   "Anger Excitatory"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdAR 
      BackColor       =   &H0000FFFF&
      Caption         =   "Anger Retaliatory"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdPR 
      BackColor       =   &H0000FFFF&
      Caption         =   "Power Reassurance"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdPA 
      BackColor       =   &H0000FFFF&
      Caption         =   "Power Assertive"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdReview 
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
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2295
   End
   Begin VB.PictureBox picface 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   6120
      Picture         =   "frmProfileRape.frx":0000
      ScaleHeight     =   4935
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   2040
      Width           =   3975
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
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   10935
   End
End
Attribute VB_Name = "frmProfileRape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form uses command buttons so the user can select what they
'believe the offenders real typology to be.

Private Sub cmdAE_Click()
    rape1answer = "four" 'this allows the user to click on a comand button to choose their profile
    frmProfileRape.Hide 'hides the current form
    frmRape1Solve.Show 'takes user to final form
End Sub

Private Sub cmdAR_Click()
    rape1answer = "three" 'this allows the user to click on a comand button to choose their profile
    frmProfileRape.Hide 'hides the current form
    frmRape1Solve.Show 'takes user to final form
End Sub


Private Sub cmdPA_Click()
    rape1answer = "one" 'this allows the user to click on a comand button to choose their profile
    frmProfileRape.Hide 'hides the current form
    frmRape1Solve.Show 'takes user to final form
End Sub

Private Sub cmdPR_Click()
    rape1answer = "two" 'this allows the user to click on a comand button to choose their profile
    frmProfileRape.Hide 'hides the current form
    frmRape1Solve.Show 'takes user to final form
End Sub

Private Sub cmdreview_Click()
'Allows for review of the case
    frmreviewcase1b.Show
End Sub

