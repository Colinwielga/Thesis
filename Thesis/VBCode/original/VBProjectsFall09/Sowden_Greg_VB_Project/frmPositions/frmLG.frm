VERSION 5.00
Begin VB.Form frmLG 
   BackColor       =   &H0000C000&
   Caption         =   "Left Guard"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label lblLG 
      BackColor       =   &H0000C000&
      Caption         =   $"frmLG.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmLG.Hide
    frmOLine.Show
    
End Sub
