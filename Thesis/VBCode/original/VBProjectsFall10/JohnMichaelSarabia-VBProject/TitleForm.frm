VERSION 5.00
Begin VB.Form frmTitlePage 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   11355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17850
   LinkTopic       =   "Form1"
   ScaleHeight     =   11355
   ScaleWidth      =   17850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Click here to quit Swim Manager 2010!"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5520
      TabIndex        =   3
      Top             =   8880
      Width           =   4455
   End
   Begin VB.CommandButton cmdTestForm 
      Caption         =   "Will your swimmer qualify for Division 3 Nationals?"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   2280
      TabIndex        =   2
      Top             =   3720
      Width           =   5175
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "What was your swimmer's average time per 50 yards?"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   9120
      TabIndex        =   1
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label lblTitle 
      Caption         =   "Welcome to swim manager 2010!!! What whould you like to do?"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   8295
   End
End
Attribute VB_Name = "frmTitlePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAverage_Click()
    frmTitlePage.Hide
    frmAverages.Show
    
End Sub
    
    
Private Sub cmdEnterTimes_Click()
    frmTitlePage.Hide
    frmTimeEvent.Show
End Sub

Private Sub cmdNationals_Click()
    frmTitlePage.Hide
    frmNationals.Show
End Sub

Private Sub cmdQuit_Click()

End

End Sub

Private Sub cmdTestForm_Click()
    frmTitlePage.Hide
    frmTestForm.Show
    
End Sub
