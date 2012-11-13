VERSION 5.00
Begin VB.Form frmPoints 
   BackColor       =   &H0080FF80&
   Caption         =   "Look Up"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBulger 
      Caption         =   "Bulger"
      Height          =   735
      Left            =   5040
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   3360
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox picOutput 
      Height          =   3255
      Left            =   360
      ScaleHeight     =   3195
      ScaleWidth      =   7275
      TabIndex        =   3
      Top             =   1800
      Width           =   7335
   End
   Begin VB.CommandButton cmdBinLaden 
      Caption         =   "Bin Laden"
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Partol Car"
      Height          =   735
      Left            =   5880
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "frmPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBinLaden_Click()
    Dim Text As String
    Open App.Path & "\BinLaden.txt" For Input As #1
    Do Until EOF(1)
        Input #1, Text
    Loop
    picOutput.Print Text
    Close #1
    
End Sub

Private Sub cmdBulger_Click()
    Dim Text2 As String
    Open App.Path & "\Bulger.txt" For Input As #2
    Do Until EOF(2)
        Input #2, Text2
    Loop
    picOutput.Print Text2
    Close #2
    
End Sub

Private Sub cmdClear_Click()
    picOutput.Cls
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmPatrolCar.Show
    frmPoints.Hide
    
End Sub
