VERSION 5.00
Begin VB.Form frmRB 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Running Back"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack8 
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
      Left            =   3360
      TabIndex        =   1
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmRB.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9855
   End
End
Attribute VB_Name = "frmRB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack8_Click()
    frmRB.Hide
    frmLearn.Show
    
End Sub
