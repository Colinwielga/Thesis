VERSION 5.00
Begin VB.Form frmEnd 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF80FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   4335
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Grand Total For Flight and Hotel"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFC0&
      Height          =   7695
      Left            =   5040
      ScaleHeight     =   7635
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   120
      Picture         =   "frmEnd.frx":0000
      Top             =   240
      Width           =   4830
   End
   Begin VB.Image Image2 
      Height          =   2565
      Left            =   120
      Picture         =   "frmEnd.frx":286DA
      Top             =   5040
      Width           =   4830
   End
End
Attribute VB_Name = "frmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Final page'
'This page displays the totals i had running'
'it adds up the flight price and hot price to get my running total'
'Blake Bauer'
'October 18th 2009'

Private Sub cmdEnd_Click()
'This is my massive print box where i am printing the totals'
    RunningTotal = TabTotal + FlightPrice + CarPrice
    picResults.Print "The Grand Total"
    picResults.Print "***************************************************"
    picResults.Print "Hotel Price"; , , ; FormatCurrency(TabTotal, 2)
    picResults.Print "Flight Price"; , , ; FormatCurrency(FlightPrice, 2)
    picResults.Print "Car Rental"; , , ; FormatCurrency(CarPrice, 2)
    picResults.Print "***************************************************"
    picResults.Print "Total Price"; , , ; FormatCurrency(RunningTotal, 2)
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print
    picResults.Print "Have A Great Trip!"
End Sub
'Quit Button'
Private Sub cmdQuit_Click()
    End
End Sub
