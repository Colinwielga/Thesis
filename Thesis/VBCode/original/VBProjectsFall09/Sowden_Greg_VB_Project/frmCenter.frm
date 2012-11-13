VERSION 5.00
Begin VB.Form frmCenter 
   BackColor       =   &H000000C0&
   Caption         =   "Center"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack9 
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   $"frmCenter.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   Center
'   Greg Sowden
'   10/10/09
'   This subroutine shows the user the information about the particular position stated on the button.

Private Sub cmdBack9_Click()
    frmCenter.Hide
    frmOLine.Show
    
End Sub
