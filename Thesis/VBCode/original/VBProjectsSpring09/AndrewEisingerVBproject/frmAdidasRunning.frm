VERSION 5.00
Begin VB.Form frmAdidasRunning 
   BackColor       =   &H0000FFFF&
   Caption         =   "AdidasRunning"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17130
   LinkTopic       =   "Form1"
   Picture         =   "frmAdidasRunning.frx":0000
   ScaleHeight     =   9780
   ScaleWidth      =   17130
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   6015
      Left            =   10320
      ScaleHeight     =   5955
      ScaleWidth      =   6315
      TabIndex        =   5
      Top             =   720
      Width           =   6375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      Height          =   1695
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Go Back to Store Home"
      Height          =   1695
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back to Adidas Store"
      Height          =   1695
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H0000FFFF&
      Caption         =   "Input"
      Height          =   1695
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H000080FF&
      Caption         =   "Read Data"
      Height          =   1695
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "frmAdidasRunning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' AdidasRunning
' Andrew Eisinger
' 3/17/09
'This program reads a file
'This program then based on the file lets the user enter a price and it matches it to an item
Dim RunningItems(1 To 250) As String, RunningCosts(1 To 250) As Single
Dim CTR As Single


Private Sub cmdBack_Click()
frmStoreHome.Show
frmAdidasRunning.Hide
End Sub

Private Sub cmdGoBack_Click()
frmAdidas1.Show
frmAdidasRunning.Hide
End Sub

Private Sub cmdInput_Click()
Dim Price As Single, Pos As Single, Found As Boolean
Pos = 1
Price = InputBox("Please Enter a maximum price you will pay", "Price")
For Pos = 1 To CTR
If Price >= RunningCosts(Pos) Then
    picResults.Print "With a price of "; FormatCurrency(Price); " you can buy "; RunningItems(Pos)
End If
Next Pos


End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()


Open App.Path & "\AdidasRunning.txt" For Input As #1
CTR = 0
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, RunningItems(CTR), RunningCosts(CTR)
Loop
Close #1

End Sub
