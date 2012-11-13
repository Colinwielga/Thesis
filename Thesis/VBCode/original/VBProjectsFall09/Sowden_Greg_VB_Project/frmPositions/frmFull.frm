VERSION 5.00
Begin VB.Form frmFull 
   BackColor       =   &H0080FF80&
   Caption         =   "Fullback"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to Positions"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   $"frmFull.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   9135
   End
End
Attribute VB_Name = "frmFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack7_Click()
    frmFull.Hide
    frmLearn.Show
    
End Sub
