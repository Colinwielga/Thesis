VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Return to Main Menu"
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdSortRoster 
      Caption         =   "Numerically Sort Roster"
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   8775
      Left            =   3600
      ScaleHeight     =   8715
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton cmdReadAndPrintRoster 
      Caption         =   "Read and Print Roster"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare Global variables for this form
Dim Names(1 To 60) As String
Dim Number(1 To 60) As Integer
Dim ctr As Integer
'Returns user to main menu
Private Sub cmdMain_Click()
    frmRoster.Hide
    frmHome.Show
End Sub
'Quits the program
Private Sub cmdQuit_Click()
 End
End Sub

Private Sub cmdReadAndPrintRoster_Click()

Dim I As Integer
'Reads roster into arrays
Open App.Path & "\Roster.txt" For Input As #1
ctr = 0
    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, Number(ctr), Names(ctr)
    Loop
    Close #1
'Prints roster into picture box to be viewed
    For I = 1 To ctr
        picResults.Print Number(I); Names(I)
    Next I
End Sub
'Sorts the players by their roster number
Private Sub cmdSortRoster_Click()
picResults.Cls
Dim pass As Integer, pos As Integer, temp1 As String, temp2 As String

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Number(pos) > Number(pos + 1) Then
            temp1 = Number(pos)
            Number(pos) = Number(pos + 1)
            Number(pos + 1) = temp1
            
            temp2 = Names(pos)
            Names(pos) = Names(pos + 1)
            Names(pos + 1) = temp2
            
        End If
    Next pos
Next pass
Dim I As Integer
For I = 1 To ctr
    picResults.Print Number(I); Names(I)
Next I

cmdReadAndPrintRoster.Enabled = False
End Sub
