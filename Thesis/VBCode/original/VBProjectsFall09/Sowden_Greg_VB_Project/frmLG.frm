VERSION 5.00
Begin VB.Form frmLG 
   Caption         =   "Left Guard"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack8 
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
      TabIndex        =   1
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lblLG 
      Caption         =   $"frmLG.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   LG
'   Greg Sowden
'   10/10/09
'   This subroutine shows the user the information about the particular position stated on the button.

Private Sub cmdBack8_Click()
    frmLG.Hide
    frmOLine.Show
    
End Sub
