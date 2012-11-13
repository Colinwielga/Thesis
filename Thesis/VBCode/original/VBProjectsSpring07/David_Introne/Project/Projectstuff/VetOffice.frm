VERSION 5.00
Begin VB.Form VetOffice 
   BackColor       =   &H00004040&
   Caption         =   "Vet Office"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H00004040&
      Caption         =   "Welcome to the vet's office! To begin, click on me."
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   10500
      Left            =   480
      Picture         =   "VetOffice.frx":0000
      Top             =   0
      Width           =   9300
   End
End
Attribute VB_Name = "VetOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Image1_Click()
VetVisit.Show ' goes to next form
VetOffice.Hide
End Sub
