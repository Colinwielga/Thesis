VERSION 5.00
Begin VB.Form frmTraffic 
   BackColor       =   &H80000000&
   Caption         =   "Traffic Sitation"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   Picture         =   "frmTraffic.frx":0000
   ScaleHeight     =   7155
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   6960
      TabIndex        =   8
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return To Patrol Car"
      Height          =   615
      Left            =   6960
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdResult 
      BackColor       =   &H80000004&
      Caption         =   "Submit"
      Height          =   615
      Left            =   3600
      MaskColor       =   &H80000004&
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000009&
      Height          =   1575
      Left            =   2400
      ScaleHeight     =   1515
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   4800
      Width           =   4215
   End
   Begin VB.TextBox txtSpeed 
      Height          =   525
      Left            =   4680
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtLimit 
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblSpeed 
      BackColor       =   &H80000005&
      Caption         =   "Speed Traveling"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblLimit 
      BackColor       =   &H80000005&
      Caption         =   "Speed Limit"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmTraffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Patrol Car
' frmTraffic
' Kevin Conlon
' 11/4/08
' write speading tickets

Private Sub cmdClear_Click()
    picOutput.Cls
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdResult_Click()
    Dim SpeedLimit As Integer
    Dim SpeedTraveling As Integer
    Dim SpeedOver As Integer
    SpeedLimit = txtLimit.Text
    SpeedTraveling = txtSpeed.Text
    
    SpeedOver = SpeedTraveling - SpeedLimit
    
    picOutput.Print "Vehicle traveling "; SpeedOver; " mph over posted limit."
    
    Select Case SpeedOver
        Case Is > 25
            picOutput.Print "Fine is $225"
            picOutput.Print "Suspend Drivers License"
        Case 20 To 24
            picOutput.Print "Fine is $175"
        Case 15 To 19
            picOutput.Print "Fine is $150"
        Case 10 To 14
            picOutput.Print "Fine is $100"
        Case 5 To 9
            picOutput.Print "Fine is $75"
        Case 1 To 5
            picOutput.Print "Issue Warning"
    End Select
End Sub

Private Sub cmdReturn_Click()
    frmPatrolCar.Show
    frmTraffic.Hide
    
End Sub
