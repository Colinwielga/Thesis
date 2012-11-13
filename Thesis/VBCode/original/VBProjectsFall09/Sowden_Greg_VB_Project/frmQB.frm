VERSION 5.00
Begin VB.Form frmQB 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Quarterback"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack11 
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label lblQB 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmQB.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmQB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   QB
'   Greg Sowden
'   10/10/09
'   This subroutine shows the user the information about the particular position stated on the button.

Private Sub cmdBack11_Click()
    frmQB.Hide
    frmLearn.Show
    
End Sub
