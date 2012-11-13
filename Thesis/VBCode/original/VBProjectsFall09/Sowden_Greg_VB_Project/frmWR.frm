VERSION 5.00
Begin VB.Form frmWR 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Wide Reciever"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmWR.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "frmWR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   WR
'   Greg Sowden
'   10/10/09
'   This subroutine shows the user the information about the particular position stated on the button.
Private Sub Command1_Click()
    frmWR.Hide
    frmLearn.Show
    
End Sub
