VERSION 5.00
Begin VB.Form frmSpain 
   BackColor       =   &H80000012&
   Caption         =   "Spain By:Jerome D'Alessandro"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "GO BACK TO COUNTRY SELECTION"
      Height          =   3615
      Left            =   9000
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox picMestalla 
      Height          =   3135
      Left            =   7800
      Picture         =   "frmSpain.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   3435
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picCampNou 
      Height          =   3135
      Left            =   3960
      Picture         =   "frmSpain.frx":3D7B
      ScaleHeight     =   3075
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picBernabeu 
      Height          =   3135
      Left            =   120
      Picture         =   "frmSpain.frx":8025
      ScaleHeight     =   3075
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   480
      ScaleHeight     =   2475
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   1320
      Width           =   7935
   End
   Begin VB.CommandButton cmdFIND 
      Caption         =   "***WHICH STADIUMS CAN I ATTEND?***"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
   Begin VB.Label lblMestalla 
      BackColor       =   &H80000012&
      Caption         =   "Estadio Mestalla"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCampNou 
      BackColor       =   &H80000012&
      Caption         =   "Camp Nou"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblBernabeu 
      BackColor       =   &H80000012&
      Caption         =   "Bernabeu"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmSpain"
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
    frmSpain.Hide
    frmSTART.Show
End Sub
Private Sub cmdFIND_Click()
    picResults.Cls
    picBernabeu.Visible = False
    lblBernabeu.Visible = False
    picCampNou.Visible = False
    lblCampNou.Visible = False
    picMestalla.Visible = False
    lblMestalla.Visible = False
        
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
            picCampNou.Visible = True
            lblCampNou.Visible = True
        End If
        If Price >= 30 Then
            picBernabeu.Visible = True
            lblBernabeu.Visible = True
        End If
        If Price >= 35 Then
            picMestalla.Visible = True
            lblMestalla.Visible = True
        End If
End Sub
Private Sub Form_Load()

    Open App.Path & "\spain.txt" For Input As #1
        For I = 1 To 3
           Input #1, nameArray(I), capacityArray(I), priceArray(I)
        Next I
    Close #1
        
End Sub

