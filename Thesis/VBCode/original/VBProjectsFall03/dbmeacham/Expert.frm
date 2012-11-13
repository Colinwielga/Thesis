VERSION 5.00
Begin VB.Form Expert 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBoots 
      Caption         =   "Select Boots"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdBinding 
      Caption         =   "Select Bindings"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdSki 
      Caption         =   "Select Skis"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Expert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBinding_Click()
Expert.Hide
ExpertBinding.Show

End Sub

Private Sub cmdBoots_Click()
Expert.Hide
ExpertBoot.Show
End Sub

Private Sub cmdSki_Click()
Expert.Hide
ExpertSki.Show

End Sub
