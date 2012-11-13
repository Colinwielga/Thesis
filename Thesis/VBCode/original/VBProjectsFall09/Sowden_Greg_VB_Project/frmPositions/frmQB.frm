VERSION 5.00
Begin VB.Form frmQB 
   BackColor       =   &H000000FF&
   Caption         =   "Quarterback"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Modern No. 20"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to Positions"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   $"frmQB.frx":0000
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "frmQB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack6_Click()
    frmQB.Hide
    frmLearn.Show
    
End Sub
