VERSION 5.00
Begin VB.Form frmRG 
   BackColor       =   &H0000C000&
   Caption         =   "Right Guard"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack2 
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
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblLG 
      BackColor       =   &H0000C000&
      Caption         =   $"frmRG.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6975
   End
End
Attribute VB_Name = "frmRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack2_Click()
    frmRG.Hide
    frmOLine.Show
    
End Sub
