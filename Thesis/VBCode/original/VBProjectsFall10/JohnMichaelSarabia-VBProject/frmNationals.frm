VERSION 5.00
Begin VB.Form frmNationals 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   12195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18930
   LinkTopic       =   "Form1"
   ScaleHeight     =   12195
   ScaleWidth      =   18930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Return to Home Page"
      Height          =   1455
      Left            =   1200
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   11175
      Left            =   5400
      ScaleHeight     =   11115
      ScaleWidth      =   13395
      TabIndex        =   1
      Top             =   840
      Width           =   13455
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click Here To Start The Program"
      Height          =   1215
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "frmNationals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pos As Integer
Dim SwimEvent(1 To 38) As String
Dim Cut(1 To 38) As String
Dim Time(1 To 38) As String
Dim CTR As Integer


Private Sub cmdClick_Click()
Dim nationalsevent As String
Dim nationalstime As Date



Open App.Path & "\miaccuttimes2010.txt" For Input As #1
CTR = 0
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, SwimEvent(CTR), Cut(CTR), Time(CTR)
    Loop
    Close #1
    
nationalsevent = InputBox("Enter Event", "Enter Event")
nationalstime = InputBox("Enter final time in seconds", "Final Time")


For Pos = 1 To CTR
    Next Pos

End Sub

Private Sub cmdQuit_Click()
frmNationals.Hide
frmTitlePage.Show

End Sub
