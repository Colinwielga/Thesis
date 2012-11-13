VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   11970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   11970
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnUserStats 
      BackColor       =   &H8000000B&
      Caption         =   "Go To User Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   5
      Top             =   8760
      Width           =   3135
   End
   Begin VB.CommandButton btnQuit1 
      BackColor       =   &H8000000B&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11520
      TabIndex        =   4
      Top             =   10320
      Width           =   2415
   End
   Begin VB.CommandButton btnQuarterBackFrm 
      BackColor       =   &H8000000B&
      Caption         =   "Go To Quarter Back Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CommandButton btnReceiverFrm 
      BackColor       =   &H8000000B&
      Caption         =   "Go To Wide Receiver Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   3135
   End
   Begin VB.CommandButton btnRunningBackFrm 
      BackColor       =   &H8000000B&
      Caption         =   "Go To Running Back Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "NFL Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13695
   End
   Begin VB.Image imgNFL 
      Height          =   11625
      Left            =   3480
      Picture         =   "NFLProject.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnQuarterBackFrm_Click()
    frmQuarterBack.Show 'This button will hide the current form, "start", and show the Quarterback form
    frmStart.Hide
End Sub

Private Sub btnQuit1_Click()
    End 'This button exits out of the program when pushed
End Sub

Private Sub btnReceiverFrm_Click()
    frmWideReciever.Show  'This button will hide the current form, "start", and show the WideReciever form
    frmStart.Hide
End Sub

Private Sub btnRunningBackFrm_Click()
    frmRunningBack.Show  'This button will hide the current form, "start", and show the Runningback form
    frmStart.Hide
End Sub

Private Sub btnUserStats_Click() 'This button will hide current form and show the form where users can enter their own stats
    frmEnterUserStats.Show
    frmStart.Hide
End Sub
