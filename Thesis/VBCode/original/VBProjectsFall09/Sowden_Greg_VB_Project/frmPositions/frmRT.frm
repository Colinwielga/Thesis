VERSION 5.00
Begin VB.Form frmRT 
   BackColor       =   &H0000C000&
   Caption         =   "Right Tackle"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to the O-Line"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label lblLT 
      BackColor       =   &H0000C000&
      Caption         =   $"frmRT.frx":0000
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
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
End
Attribute VB_Name = "frmRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack4_Click()
    frmRT.Hide
    frmOLine.Show
    
End Sub
