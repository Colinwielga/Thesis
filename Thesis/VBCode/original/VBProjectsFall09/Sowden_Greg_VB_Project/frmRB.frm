VERSION 5.00
Begin VB.Form frmRB 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Running Back"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback12 
      BackColor       =   &H00404040&
      Caption         =   "Go Back to Positions"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmRB.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmRB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   RB
'   Greg Sowden
'   10/10/09
'   This subroutine shows the user the information about the particular position stated on the button.

Private Sub cmdback12_Click()
    frmRB.Hide
    frmLearn.Show
End Sub
