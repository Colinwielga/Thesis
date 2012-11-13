VERSION 5.00
Begin VB.Form frmEndForm 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF80&
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFF80&
      Caption         =   "Quit Program"
      Height          =   855
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdStarOver 
      BackColor       =   &H00FFFF80&
      Caption         =   "Return to beginning"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   1455
      Left            =   2760
      ScaleHeight     =   1395
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FFFF80&
      Caption         =   "Find your grand total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblSummer 
      BackStyle       =   0  'Transparent
      Caption         =   "We hope to see you this summer!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   4575
   End
End
Attribute VB_Name = "frmEndForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Big Sky Resort
'frmEntryForm
'Ryan Hoffmann and Jamison Murphy
'Written on March 19, 2009
'This Form finds the grand total of what it would cost under
'the users desired conditions
Option Explicit
'This command goes back to previous form
Private Sub cmdBack_Click()
    frmAvailabilityForm.Show
    frmEndForm.Hide
End Sub

'This command terminates program
Private Sub cmdQuit_Click()
    End
End Sub

'This command brings user back to the start screen
Private Sub cmdStarOver_Click()
    frmEntryForm.Show
    frmEndForm.Hide
End Sub

'This command finds the grand total of lodging and activites entered earlier
Private Sub cmdTotal_Click()
    picResults.Cls
    picResults.Print "Lodging:"; FormatCurrency(LodgingTotal)
    picResults.Print "Activities:"; FormatCurrency(ActivitiesTotal)
    picResults.Print "Sales Tax: 7%"
    picResults.Print "************************"
    picResults.Print "Your dream vacation will cost"
    picResults.Print "an Estimated "; FormatCurrency((LodgingTotal + ActivitiesTotal) * 1.07); "."
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub

