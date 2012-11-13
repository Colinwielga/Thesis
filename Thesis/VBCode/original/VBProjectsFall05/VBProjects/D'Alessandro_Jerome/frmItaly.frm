VERSION 5.00
Begin VB.Form frmItaly 
   BackColor       =   &H80000007&
   Caption         =   "Italy By:Jerome D'Alessandro"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   ForeColor       =   &H80000006&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "GO BACK TO COUNTRY SELECTION"
      Height          =   3015
      Left            =   8880
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picDelleAlpi 
      Height          =   3015
      Left            =   7320
      Picture         =   "frmItaly.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picOlimpico 
      Height          =   3015
      Left            =   3600
      Picture         =   "frmItaly.frx":75F2
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picSanSiro 
      Height          =   3015
      Left            =   120
      Picture         =   "frmItaly.frx":CB67
      ScaleHeight     =   2955
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   8235
      TabIndex        =   1
      Top             =   960
      Width           =   8295
   End
   Begin VB.CommandButton cmdFIND 
      BackColor       =   &H8000000E&
      Caption         =   "***WHICH STADIUMS CAN I ATTEND?***"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label lblDelleAlpi 
      BackColor       =   &H80000012&
      Caption         =   "Delle Alpi"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblOlimpico 
      BackColor       =   &H80000012&
      Caption         =   "Stadio Olimpico"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblSanSiro 
      BackColor       =   &H80000008&
      Caption         =   "Stadio San Siro"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmItaly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Price As Double
    Dim priceArray(1 To 5) As Double
    Dim nameArray(1 To 5) As String
    Dim capacityArray(1 To 5) As Double
    Dim I As Double
    Dim J As Double
Private Sub cmdBack_Click()
    frmItaly.Hide
    frmSTART.Show
End Sub
Private Sub cmdFIND_Click()
    picResults.Cls
    picSanSiro.Visible = False
    lblSanSiro.Visible = False
    picOlimpico.Visible = False
    lblOlimpico.Visible = False
    picDelleAlpi.Visible = False
    lblDelleAlpi.Visible = False
        
    Price = InputBox("HOW MUCH CAN YOU SPEND IN EUROS?", "MONEY AVAILABLE")
    picResults.Cls
    picResults.Print "Stadium / Team", ,
    picResults.Print "Capacity", ,
    picResults.Print "Price per Ticket(euros)"
    picResults.Print
        
    For J = 1 To 3
        If Price >= priceArray(J) Then
            picResults.Print nameArray(J), ,
            picResults.Print capacityArray(J), ,
            picResults.Print priceArray(J)
            picResults.Print " "
        End If
    Next J
        If Price >= 25 Then
            picOlimpico.Visible = True
            lblOlimpico.Visible = True
        End If
        If Price >= 30 Then
            picSanSiro.Visible = True
            lblSanSiro.Visible = True
        End If
        If Price >= 35 Then
            picDelleAlpi.Visible = True
            lblDelleAlpi.Visible = True
        End If
End Sub
Private Sub Form_Load()
    
    Open App.Path & "\italy.txt" For Input As #1
        For I = 1 To 3
           Input #1, nameArray(I), capacityArray(I), priceArray(I)
        Next I
    Close #1
        
End Sub

