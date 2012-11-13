VERSION 5.00
Begin VB.Form frmReebokAdults 
   BackColor       =   &H00C0C000&
   Caption         =   "ReebokAdults"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   Picture         =   "frmReebokAdults.frx":0000
   ScaleHeight     =   8400
   ScaleWidth      =   15690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sort Alphabetically"
      Height          =   1335
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H000080FF&
      Caption         =   "Go Back To Store Home"
      Height          =   1335
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdGoRe 
      BackColor       =   &H0000FF00&
      Caption         =   "Go Back To Reebok Home"
      Height          =   1335
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   7335
      Left            =   10080
      ScaleHeight     =   7275
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H000080FF&
      Caption         =   "Read"
      Height          =   1335
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmReebokAdults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' AthleticStore
' ReebokAdults
' Andrew Eisinger
' 3/22/09
'This program loads names and prints them
'Then sorts them in alphabetical order
Dim Athletes(1 To 250) As String

Private Sub cmdGo_Click()
frmStoreHome.Show
frmReebokAdults.Hide
End Sub

Private Sub cmdGoRe_Click()
frmReebok1.Show
frmReebokAdults.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()
picResults.Cls
Open App.Path & "\ReebokAdultsAthletes.txt" For Input As #1
CTR = 0
picResults.Print "Reebok Athlete's Names"
picResults.Print "************************************************************"
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Athletes(CTR)
    picResults.Print Athletes(CTR)
Loop
Close #1
End Sub



Private Sub cmdSort_Click()
'This button sorts the arrays by the first names of the products
Dim J As Single, TempAthletes As String
    'Clears the previous results
    picResults.Cls
    
    picResults.Print "Baseball Names"
    picResults.Print "**************************************************************************"

    
    'Code to sort a parralel array by the first name of the player
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Athletes(Pos) > Athletes(Pos + 1) Then
                TempAthletes = Athletes(Pos)
                Athletes(Pos) = Athletes(Pos + 1)
                Athletes(Pos + 1) = TempAthletes

            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print Athletes(J)
    Next
End Sub

