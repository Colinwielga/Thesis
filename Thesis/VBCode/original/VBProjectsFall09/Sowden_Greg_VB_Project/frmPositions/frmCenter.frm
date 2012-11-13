VERSION 5.00
Begin VB.Form frmCenter 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack3 
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
      Height          =   1095
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   $"frmCenter.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
   End
End
Attribute VB_Name = "frmCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack3_Click()
    frmCenter.Hide
    frmOLine.Show
End Sub
