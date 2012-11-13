VERSION 5.00
Begin VB.Form frmreviewcase1b 
   BackColor       =   &H00000000&
   Caption         =   "Review Case 1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picResult 
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
      Height          =   7095
      Left            =   240
      ScaleHeight     =   7095
      ScaleWidth      =   13575
      TabIndex        =   1
      Top             =   960
      Width           =   13575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   3495
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Case Review"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5520
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmreviewcase1b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This is simply a extra form built so the user can click on a button
'and then re-read the case to pick out more detail that are needed in order
'to correctly profile the offender.


Private Sub cmdReturn_Click()
'hides the form so the user can once again access the other page
    frmreviewcase1b.Hide
End Sub


Private Sub Form_Activate()
'Displays the results as soon as the form is activated
Dim ctr As Integer 'declare variables for the button
picResult.Cls 'clear out the picture box
For ctr = 1 To 17 'range i want to program to print
    picResult.Print CaseFile(ctr) 'Print the lines
Next ctr 'until the end of the file
End Sub

