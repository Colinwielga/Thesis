VERSION 5.00
Begin VB.Form FormSaves4 
   BackColor       =   &H80000007&
   Caption         =   "Save"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "It seems that you've already saved a game before. Do you want to cover it?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "FormSaves4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Open App.Path & "\Saves\Save.txt" For Output As #3
    stage = 4
    Write #3, stage
    MsgBox "Game saved."
    FormSaves4.Hide
    Close #3
End Sub

Private Sub Command2_Click()
    FormSaves4.Hide
End Sub
