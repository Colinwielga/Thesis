VERSION 5.00
Begin VB.Form namePup 
   BackColor       =   &H00800080&
   Caption         =   "What's it's name?"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "<--Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save Name and Next-->"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   5295
   End
   Begin VB.TextBox Txtname 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1695
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   5295
   End
   Begin VB.Label namepuplbl 
      BackColor       =   &H00800080&
      Caption         =   "Name your pup!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   6255
   End
End
Attribute VB_Name = "namePup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Pupname = Txtname.Text
'this tests to see which puppy has been clicked on and then sends the player to the appropriet form as a result.
    Select Case puppick
        Case 11 'moves on to next case if previous did not work
             namePup.Hide
             ProShep.Show
        Case 12 'moves on to next case if previous did not work
            ProPit.Show
            namePup.Hide
        Case 13 'moves on to next case if previous did not work
            ProMtn.Show
            namePup.Hide
        Case 14
            Produch.Show
            namePup.Hide
    End Select 'ends select
    

End Sub

Private Sub Command3_Click()
namePup.Hide 'shows and hides a form
PupsPick.Show
End Sub
