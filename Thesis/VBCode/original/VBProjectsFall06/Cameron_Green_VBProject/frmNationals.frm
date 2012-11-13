VERSION 5.00
Begin VB.Form frmNationals 
   BackColor       =   &H00008000&
   Caption         =   "Division 3 Nationals Results"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Race Results Page"
      Height          =   1215
      Left            =   7200
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for Nationals Results"
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   4035
      Left            =   3600
      Picture         =   "frmNationals.frx":0000
      Top             =   1440
      Width           =   3150
   End
End
Attribute VB_Name = "frmNationals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmNationals.Hide
    frmRaceResults.Show
End Sub
