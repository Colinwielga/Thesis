VERSION 5.00
Begin VB.Form frmBiography 
   BackColor       =   &H00FFFF80&
   Caption         =   "Shaun White Bio"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Retrun to main page"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   6240
      Width           =   2655
   End
   Begin VB.PictureBox picSpin 
      Height          =   4095
      Left            =   0
      Picture         =   "frmBiography.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   7275
      TabIndex        =   1
      Top             =   2760
      Width           =   7335
   End
   Begin VB.PictureBox picTomato 
      Height          =   4695
      Left            =   7440
      Picture         =   "frmBiography.frx":1D6A1
      ScaleHeight     =   4635
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label lblBio 
      BackColor       =   &H00FFFF80&
      Caption         =   $"frmBiography.frx":21015
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmBiography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this page offers a short bio on Shaun White and a link back to the main page
Private Sub cmdReturn_Click()
    frmShaunWhite.Show
    frmBiography.Hide
End Sub
