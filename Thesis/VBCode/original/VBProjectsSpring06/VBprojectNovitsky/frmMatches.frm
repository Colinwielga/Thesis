VERSION 5.00
Begin VB.Form frmMatches 
   BackColor       =   &H0000FFFF&
   Caption         =   "Your top matches were!"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Calculations!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Image imgMatches 
      Height          =   2595
      Left            =   4080
      Picture         =   "frmMatches.frx":0000
      Top             =   1680
      Width           =   2310
   End
End
Attribute VB_Name = "frmMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click() 'returns to frmsecond
    frmMatches.Hide
    frmSecond.Show
End Sub
