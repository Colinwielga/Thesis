VERSION 5.00
Begin VB.Form frmTickets 
   BackColor       =   &H00000080&
   Caption         =   "Tickets"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H0000FFFF&
      Caption         =   "Buy Tickets Now!"
      Height          =   615
      Index           =   1
      Left            =   6360
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdDirections 
      Caption         =   "Directions"
      Height          =   615
      Index           =   0
      Left            =   3840
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdArenaFacts 
      Caption         =   "Arena Facts"
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   7155
      TabIndex        =   3
      Top             =   3360
      Width           =   7215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   600
      Picture         =   "frmTickets.frx":0000
      ScaleHeight     =   2025
      ScaleWidth      =   2985
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00000080&
      Caption         =   "Home"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label lblMariucci 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Mariucci Arena"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3210
   End
End
Attribute VB_Name = "frmTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmTickets
'Cole and John
'10/30/06
'Objective: The objective of this form is to present the user with the oprion to
'(1) view information on mariucci arena, (2) access directions to the arena, and
'(3) continue with the ticket purchasing process. The user can access information on
'mariucci arena by clicking on the appropriate command button that fills and prints
'an array.

Option Explicit

Private Sub cmdArenaFacts_Click()

Dim Facts(1 To 11) As String
Dim Pos As Integer


    picResults.Cls
    
    Open App.Path & "\ArenaFacts.txt" For Input As #1       'opens text file
    Pos = 0
    
        Do Until EOF(1)                 'reads text file until end of list
            Pos = Pos + 1
            Input #1, Facts(Pos)        'inputs array
        Loop
     Close #1
     
    For Pos = 1 To 11               'prints array
        picResults.Print Facts(Pos)
    Next Pos
    
End Sub

Private Sub cmdBuy_Click(Index As Integer)
    frmTickets.Visible = False
    frmTicketSales.Visible = True
End Sub

Private Sub cmdDirections_Click(Index As Integer)

Dim Directions(1 To 14) As String
Dim Pos As Integer

    picResults.Cls
    
       
    Open App.Path & "\directions.txt" For Input As #1
    Pos = 0
    
        Do Until EOF(1)
            Pos = Pos + 1
            Input #1, Directions(Pos)
        Loop
     Close #1
     
    For Pos = 1 To 14
        picResults.Print Directions(Pos)
    Next Pos
       
End Sub

Private Sub cmdHome_Click()
    frmTickets.Visible = False
    frmMain.Visible = True
End Sub

