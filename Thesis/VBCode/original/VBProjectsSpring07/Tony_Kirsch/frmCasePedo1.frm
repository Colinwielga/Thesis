VERSION 5.00
Begin VB.Form frmCasePedo1 
   BackColor       =   &H00000000&
   Caption         =   "Pedophile Case 1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdbegin 
      BackColor       =   &H0000FF00&
      Caption         =   "Click to Begin Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   2895
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   2415
   End
   Begin VB.PictureBox picPedo1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   480
      ScaleHeight     =   6375
      ScaleWidth      =   13575
      TabIndex        =   0
      Top             =   240
      Width           =   13575
   End
   Begin VB.Label lblinstruction 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmCasePedo1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   5640
      TabIndex        =   3
      Top             =   8280
      Width           =   4095
   End
End
Attribute VB_Name = "frmCasePedo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is designed to allow the user to read the case file.


Private Sub cmdBack_Click()
'This command button takes the user back to the case files
    frmCasePedo1.Hide
    frmCasefiles.Show
    
End Sub

Private Sub cmdbegin_Click()
'This button takes the user to the profiling process.
'It leads to the molester type form
    frmCasePedo1.Hide
    frmProfile3.Show
End Sub

'I made command on the form itself. This makes my text appear from my file into
'my picture box as soon as the the particular form is activated
Private Sub Form_Activate()
'this declares my ctr because i use it to tell the program where
'to stop and start reading my file.
Dim ctr As Integer
    picPedo1.Cls 'Clears the picture box so i don't have any old data in there
For ctr = 37 To 55 'The part of my txt file i want to read when i activate the form.
    picPedo1.Print CaseFile(ctr) 'Print it out
Next ctr 'Until the ctr has run the length of the program
End Sub
