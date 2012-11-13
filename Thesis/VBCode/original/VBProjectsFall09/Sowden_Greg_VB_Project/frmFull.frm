VERSION 5.00
Begin VB.Form frmFull 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Fullback"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback14 
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmFull.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   Full
'   Greg Sowden
'   10/10/09
'   This subroutine shows the user the information about the particular position stated on the button.

Private Sub cmdback14_Click()
   frmFull.Hide
    frmLearn.Show
    
End Sub
