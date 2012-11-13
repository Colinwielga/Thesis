VERSION 5.00
Begin VB.Form frmPatrolCar 
   BackColor       =   &H80000000&
   Caption         =   "Patrol Car"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   Picture         =   "frmPatrolCar.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDay 
      BackColor       =   &H80000000&
      Caption         =   "Day Mode"
      Height          =   735
      Left            =   7680
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdNight 
      Caption         =   "Night Mode"
      Height          =   735
      Left            =   7680
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdWanted 
      Caption         =   "       Wanted List        10-29  "
      Height          =   735
      Left            =   7680
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdPoints 
      Caption         =   "Look Up License  10-26"
      Height          =   735
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdTraffic 
      Caption         =   "Make Traffic Stop 11-95"
      Height          =   735
      Left            =   7680
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "What cop are you?"
      Height          =   735
      Left            =   7680
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave Squad Car 10-7"
      Height          =   735
      Left            =   7680
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
   End
End
Attribute VB_Name = "frmPatrolCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Patrol Car
' frmpatrolcar
' Kevin Conlon
' 11/4/08
' to be the main page of program

Private Sub cmdDay_Click()
    frmPatrolCar.BackColor = &H8000000F
    cmdNight.Visible = True
    cmdDay.Visible = False
End Sub

Private Sub cmdNight_Click()
    frmPatrolCar.BackColor = &H0&
    cmdNight.Visible = False
    cmdDay.Visible = True
End Sub

Private Sub cmdPoints_Click()
    frmPatrolCar.Hide
    frmPoints.Show
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdQuiz_Click()
    frmQuiz.Show
    frmPatrolCar.Hide
End Sub

Private Sub cmdTraffic_Click()
    ' Go traffic form
    frmPatrolCar.Hide
    frmTraffic.Show
    
End Sub

Private Sub cmdWanted_Click()
    ' Go Wanted form
    frmPatrolCar.Hide
    frmWanted.Show
    
End Sub
