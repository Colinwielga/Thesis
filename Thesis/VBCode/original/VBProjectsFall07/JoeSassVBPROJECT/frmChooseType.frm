VERSION 5.00
Begin VB.Form frmChooseType 
   Caption         =   "What format did you use?"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAPA 
      Caption         =   "APA"
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdMLA 
      Caption         =   "MLA"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblWhatType 
      Alignment       =   2  'Center
      Caption         =   "What type of format did you use? Please choose below:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmChooseType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form asks the user which type of bibliography they are loading, which determines how it is formated on the frmLoaded form

Private Sub cmdAPA_Click()
    'Sets the type as APA and then hides this form
    MLA = False
    frmChooseType.Hide
End Sub

Private Sub cmdMLA_Click()
    'sets the type to MLA and then hides this form
    MLA = True
    frmChooseType.Hide
End Sub
