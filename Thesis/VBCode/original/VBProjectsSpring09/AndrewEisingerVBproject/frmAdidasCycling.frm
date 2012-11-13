VERSION 5.00
Begin VB.Form frmAdidasCycling 
   BackColor       =   &H8000000C&
   Caption         =   "AdidasCycling"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   Picture         =   "frmAdidasCycling.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   6615
      Left            =   6000
      ScaleHeight     =   6555
      ScaleWidth      =   5355
      TabIndex        =   5
      Top             =   600
      Width           =   5415
   End
   Begin VB.CommandButton cmdGoStore 
      BackColor       =   &H000000FF&
      Caption         =   "Go to Store Home"
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H0000FF00&
      Caption         =   "Go to Adidas Store"
      Height          =   1335
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H00404040&
      Caption         =   "Input"
      Height          =   1335
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H8000000C&
      Caption         =   "Read Data"
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdidasCycling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' AdidasCycling
' Andrew Eisinger
' 3/17/09
'This program reads a file
'This program then based on the file lets the user enter a price and it matches it to an item
Dim CyclingItems(1 To 250) As String, CyclingCosts(1 To 250) As Single, CTR As Single


Private Sub cmdGoBack_Click()
frmAdidas1.Show
frmAdidasCycling.Hide
End Sub

Private Sub cmdGoStore_Click()
frmStoreHome.Show
frmAdidasCycling.Hide
End Sub

Private Sub cmdInput_Click()
Dim Price As Single

Pos = 1
Price = InputBox("Please Enter a maximum Price you would pay", "Price")
For Pos = 1 To CTR
If Price >= CyclingCosts(Pos) Then
    picResults.Print "With a price of "; FormatCurrency(Price); " you can buy "; CyclingItems(Pos)
   End If
Next Pos

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()

Open App.Path & "\AdidasCycling.txt" For Input As #1
CTR = 0
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, CyclingItems(CTR), CyclingCosts(CTR)
Loop
Close #1

End Sub
