VERSION 5.00
Begin VB.Form frmEngland 
   BackColor       =   &H80000008&
   Caption         =   "England By: Jerome D'Alessandro"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   ScaleHeight     =   7470
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000E&
      Caption         =   "GO BACK TO COUNTRY SELCTION"
      Height          =   3015
      Left            =   9000
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdFIND 
      Caption         =   "***WHICH STADIUMS I CAN ATTEND?***"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   8295
   End
   Begin VB.PictureBox picTrafford 
      Height          =   3015
      Left            =   7320
      Picture         =   "frmEngland.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picHighbury 
      Height          =   3015
      Left            =   3720
      Picture         =   "frmEngland.frx":4AC9
      ScaleHeight     =   2955
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picResults 
      Height          =   2175
      Left            =   360
      ScaleHeight     =   2115
      ScaleWidth      =   8235
      TabIndex        =   2
      Top             =   960
      Width           =   8295
   End
   Begin VB.PictureBox picStamford 
      Height          =   3015
      Left            =   240
      Picture         =   "frmEngland.frx":4173B
      ScaleHeight     =   2955
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblTrafford 
      BackColor       =   &H80000012&
      Caption         =   "Old Trafford"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblHighbury 
      BackColor       =   &H80000012&
      Caption         =   "Highbury"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblStamford 
      BackColor       =   &H80000012&
      Caption         =   "Stamford Bridge"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmEngland"
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
    frmEngland.Hide
    frmSTART.Show
End Sub
Private Sub cmdFIND_Click()
    picResults.Cls
    picTrafford.Visible = False
    lblTrafford.Visible = False
    picStamford.Visible = False
    lblStamford.Visible = False
    picHighbury.Visible = False
    lblHighbury.Visible = False
        
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
            picTrafford.Visible = True
            lblTrafford.Visible = True
        End If
        If Price >= 30 Then
            picStamford.Visible = True
            lblStamford.Visible = True
        End If
        If Price >= 40 Then
            picHighbury.Visible = True
            lblHighbury.Visible = True
        End If
End Sub
Private Sub Form_Load()

    Open App.Path & "\england.txt" For Input As #1
        For I = 1 To 3
           Input #1, nameArray(I), capacityArray(I), priceArray(I)
        Next I
    Close #1
        
End Sub

