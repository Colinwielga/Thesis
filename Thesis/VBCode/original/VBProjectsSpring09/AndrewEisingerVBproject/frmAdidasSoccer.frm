VERSION 5.00
Begin VB.Form frmAdidasSoccer 
   BackColor       =   &H80000007&
   Caption         =   "AdidasSoccer"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   Picture         =   "frmAdidasSoccer.frx":0000
   ScaleHeight     =   8400
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5415
      Left            =   9600
      ScaleHeight     =   5355
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   1200
      Width           =   4935
   End
   Begin VB.CommandButton cmdGoBackHome 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Go Back to Store Home"
      Height          =   1455
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Quit"
      Height          =   1455
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00FF8080&
      Caption         =   "Go Back to Adidas Store"
      Height          =   1575
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H00FF8080&
      Caption         =   "Input"
      Height          =   1575
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FF8080&
      Caption         =   "Read Data"
      Height          =   1575
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmAdidasSoccer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' AdidasSoccer
' Andrew Eisinger
' 3/17/09
'This program reads a file
'This program then based on the file lets the user enter a price and it matches it to an item
Dim SoccerItems(1 To 100) As String, SoccerCosts(1 To 100) As Single, CTR As Single

Private Sub cmdGoBack_Click()
frmAdidas1.Show
frmAdidasSoccer.Hide
End Sub

Private Sub cmdGoBackHome_Click()
frmStoreHome.Show
frmAdidasSoccer.Hide
End Sub

Private Sub cmdInput_Click()
Dim Price As Single

Pos = 1
Price = InputBox("Please Enter a maximum price you can pay", "Price")
For Pos = 1 To CTR
If Price >= SoccerCosts(Pos) Then
    picResults.Print "With a price of "; FormatCurrency(Price); " you can buy "; SoccerItems(Pos)
   End If
Next Pos
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()

Open App.Path & "\AdidasSoccer.txt" For Input As #1
CTR = 0
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, SoccerItems(CTR), SoccerCosts(CTR)
Loop
Close #1
End Sub

