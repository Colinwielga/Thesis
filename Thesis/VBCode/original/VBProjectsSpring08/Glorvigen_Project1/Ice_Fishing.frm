VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form6"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form6"
   ScaleHeight     =   5535
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdexit 
      Caption         =   "Back to Main Page"
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ice Fishing"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdexit_Click()
    Form1.Show
    Form6.Hide
End Sub
