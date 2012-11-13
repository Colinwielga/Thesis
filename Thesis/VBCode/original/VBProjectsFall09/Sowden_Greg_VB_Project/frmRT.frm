VERSION 5.00
Begin VB.Form frmRT 
   BackColor       =   &H00FF0000&
   Caption         =   "Right Tackle"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to O-Line"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label lblLT 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmRT.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
End
Attribute VB_Name = "frmRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   RT
'   Greg Sowden
'   10/10/09
'   This subroutine shows the user the information about the particular position stated on the button.
Private Sub cmdBack7_Click()
    frmRT.Hide
    frmOLine.Show
    
End Sub
