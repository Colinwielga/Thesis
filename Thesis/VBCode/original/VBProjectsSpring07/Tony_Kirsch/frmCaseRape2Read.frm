VERSION 5.00
Begin VB.Form frmCaseRape2Read 
   BackColor       =   &H00000000&
   Caption         =   "Rape Case 2"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H0000C000&
      Caption         =   "Press to Create your profile"
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
      Left            =   10680
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   3135
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
      Height          =   1335
      Left            =   2640
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   2295
   End
   Begin VB.PictureBox picRape2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   7215
      Left            =   1320
      ScaleHeight     =   7215
      ScaleWidth      =   10455
      TabIndex        =   0
      Top             =   240
      Width           =   10455
   End
   Begin VB.Label lblinstruction 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmCaseRape2Read.frx":0000
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
      TabIndex        =   2
      Top             =   7800
      Width           =   4095
   End
End
Attribute VB_Name = "frmCaseRape2Read"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBack_Click()
'Allows the user to go back to case file form
    frmCaseRape2Read.Hide
    frmCasefiles.Show
    
End Sub

Private Sub cmdbegin_Click()
'This allows the user to move on to the profiler stage
    frmCaseRape2Read.Hide
    frmProfile2.Show
    
End Sub

'I use the form activate function to display my data as soon as the form itself
'is activated.
Private Sub Form_Activate()
'Declare my varaible
Dim ctr As Integer

    picRape2.Cls 'Clears picture box of any data.
For ctr = 19 To 35 'the range i want my file read
    picRape2.Print CaseFile(ctr) 'Print out my range
Next ctr 'Until the end of the ctr value

End Sub

