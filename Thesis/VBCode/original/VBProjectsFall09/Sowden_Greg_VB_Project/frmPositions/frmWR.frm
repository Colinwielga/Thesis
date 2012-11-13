VERSION 5.00
Begin VB.Form frmWR 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Wide Reciever"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack5 
      BackColor       =   &H000000FF&
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
      Height          =   1215
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label lblWR 
      BackColor       =   &H00FFFFFF&
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
      Height          =   6495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmWR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack5_Click()
    frmWR.Hide
    frmLearn.Show
    
End Sub
