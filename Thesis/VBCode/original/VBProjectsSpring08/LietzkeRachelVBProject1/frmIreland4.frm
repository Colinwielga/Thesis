VERSION 5.00
Begin VB.Form frmIreland4 
   BackColor       =   &H000080FF&
   Caption         =   "Ireland 4"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   Picture         =   "frmIreland4.frx":0000
   ScaleHeight     =   4995
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Get Average Hotel Price"
      Height          =   735
      Left            =   4560
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdGoHome 
      Caption         =   "Go Home"
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtRate 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtDays 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblDays 
      Caption         =   "Input Nights Here"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblRate 
      BackColor       =   &H00FF8080&
      Caption         =   "Input Rate of Hotel Here."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   5865
      Left            =   -120
      Picture         =   "frmIreland4.frx":5C05
      Top             =   -240
      Width           =   6990
   End
End
Attribute VB_Name = "frmIreland4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Information About Ireland
'Form Name: Ireland4
'Author: Rachel Lietzke
'Date Written: March 27, 2008
'Objective: To offer the User a way to caucluate their Hotel Cost and
'To give the Average Cost from the Hotels Previously Given

Private Sub cmdAverage_Click()
Dim Average As Single
Dim City(1 To 30) As String
Dim Names(1 To 30) As String
Dim Stars(1 To 30) As Integer
Dim Cost(1 To 30) As Single
Dim CTR As Integer
Dim Sum As Single

Sum = 0
CTR = 0

Open App.Path & "\More.txt" For Input As #1

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, City(CTR), Names(CTR), Stars(CTR), Cost(CTR)
    Sum = Sum + Cost(CTR)
Loop

Average = Sum / CTR



MsgBox "Tha average price for a Hotel Room is " & FormatCurrency(Int(Average), 2) & ".", , "Average Price"

End Sub

Private Sub cmdCalculate_Click()
Dim Rate As Single
Dim Days As Single
Dim Total As Single

Rate = txtRate.Text
Days = txtDays.Text

Total = Rate * Days

MsgBox "You will pay " & FormatCurrency(Total, 2) & " for the Hotel you choose."
End Sub

Private Sub cmdGoHome_Click()
frmIreland4.Hide
frmIreland1.Show
End Sub

Private Sub Label1_Click()

End Sub
