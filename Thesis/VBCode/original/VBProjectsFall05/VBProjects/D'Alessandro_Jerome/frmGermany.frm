VERSION 5.00
Begin VB.Form frmGermany 
   BackColor       =   &H80000012&
   Caption         =   "Germany By: Jerome D'Alessandro"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "GO BACK TO COUNTRY SELECTION"
      Height          =   3735
      Left            =   8640
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox picWestfalenstadion 
      Height          =   3015
      Left            =   7560
      Picture         =   "frmGermany.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picWeserstadion 
      Height          =   3015
      Left            =   3840
      Picture         =   "frmGermany.frx":44F6
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picAllianz 
      Height          =   3015
      Left            =   120
      Picture         =   "frmGermany.frx":9C01
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picResults 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   7995
      TabIndex        =   1
      Top             =   1320
      Width           =   8055
   End
   Begin VB.CommandButton cmdFIND 
      Caption         =   "***WHICH STADIUMS CAN I ATTEND?***"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8055
   End
   Begin VB.Label lblWestfalenstadion 
      BackColor       =   &H80000012&
      Caption         =   "Westfalenstadion"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblWeserstadion 
      BackColor       =   &H80000012&
      Caption         =   "Weserstadion"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblAllianz 
      BackColor       =   &H80000012&
      Caption         =   "Allianz Arena"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmGermany"
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
    frmGermany.Hide
    frmSTART.Show
End Sub
Private Sub cmdFIND_Click()
    picResults.Cls
    picAllianz.Visible = False
    lblAllianz.Visible = False
    picWeserstadion.Visible = False
    lblWeserstadion.Visible = False
    picWestfalenstadion.Visible = False
    lblWestfalenstadion.Visible = False
        
    Price = InputBox("HOW MUCH CAN YOU SPEND IN EUROS?", "MONEY AVAILABLE")
    picResults.Cls
    picResults.Print "Stadium / Team", , ,
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
        If Price >= 30 Then
            picWestfalenstadion.Visible = True
            lblWestfalenstadion.Visible = True
        End If
        If Price >= 35 Then
            picAllianz.Visible = True
            lblAllianz.Visible = True
        End If
        If Price >= 40 Then
            picWeserstadion.Visible = True
            lblWeserstadion.Visible = True
        End If
End Sub
Private Sub Form_Load()
    
    Open App.Path & "\germany.txt" For Input As #1
        For I = 1 To 3
           Input #1, nameArray(I), capacityArray(I), priceArray(I)
        Next I
    Close #1
        
End Sub

