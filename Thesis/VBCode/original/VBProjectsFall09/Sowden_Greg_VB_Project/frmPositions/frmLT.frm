VERSION 5.00
Begin VB.Form frmLT 
   BackColor       =   &H0000C000&
   Caption         =   "Left Tackle"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to the O-Line"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label lblLT 
      BackColor       =   &H0000C000&
      Caption         =   $"frmLT.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   7335
   End
End
Attribute VB_Name = "frmLT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack1_Click()
    frmLT.Hide
    frmOLine.Show
    
End Sub
