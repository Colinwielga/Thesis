VERSION 5.00
Begin VB.Form frmCasePedo2 
   BackColor       =   &H00000000&
   Caption         =   "Pedophile case 2"
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
      Caption         =   "Press to begin making your profile"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
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
      Left            =   1200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8760
      Width           =   2655
   End
   Begin VB.PictureBox picpedo2 
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
      Height          =   8175
      Left            =   120
      ScaleHeight     =   8175
      ScaleWidth      =   15015
      TabIndex        =   0
      Top             =   120
      Width           =   15015
   End
   Begin VB.Label lblinstruction 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmCasePedo2.frx":0000
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
      Left            =   5400
      TabIndex        =   3
      Top             =   8520
      Width           =   4095
   End
End
Attribute VB_Name = "frmCasePedo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
'Takes me from my particular case file back to the list of case files
    frmCasePedo2.Hide
    frmCasefiles.Show
    
End Sub

Private Sub cmdbegin_Click()
'Takes me from my case file text to the form where i start profiling
    frmCasePedo2.Hide
    frmProfile4.Show
End Sub

'I made command on the form itself. This makes my text appear from my file into
'my picture box as soon as the the particular form is activated
Private Sub Form_Activate()
'this declares my ctr because i use it to tell the program where
'to stop and start reading my file.
Dim ctr As Integer
    picpedo2.Cls 'Clears the picture box so i don't have any old data in there
For ctr = 57 To 83 'The part of my txt file i want to read when i activate the form.
    picpedo2.Print CaseFile(ctr) 'Print it out
Next ctr 'Until the ctr has run the length of the program

End Sub
