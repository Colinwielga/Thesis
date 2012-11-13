VERSION 5.00
Begin VB.Form frmGas 
   Caption         =   "Gas Mileage"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   Picture         =   "frmGas.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMileage 
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Find number of cars with particular gas mileage"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   2895
      Left            =   1680
      ScaleHeight     =   2835
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label lblEnter 
      Caption         =   "Enter gas mileage here  =>"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
End
Attribute VB_Name = "frmGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmGas
'Author is Dan Parker
'Date written 10/18/09
'The purpose of this form is to give the user the number and name of a particular
'set of vehicles that receive or better the gas mileage gives by the user.

Private Sub cmdBack_Click()
    'brings user back to homepage
    frmGas.Hide
    frmFirst.Show
End Sub

Private Sub cmdCalc_Click()
    'dim local variables
    Dim gasMileage As Single, I As Integer, found As Boolean, car(1 To 20) As String, mileage(1 To 20) As Integer, ctr As Integer, whoctr As Integer
    
    picResults.Cls
    picResults.Print "Car"; Tab(38); "Average miles per gallon"
    picResults.Print "***********************************************************************************"
    gasMileage = txtMileage.Text 'set gas mileage given by user equal to the number given in the text box
    whoctr = 0
    found = False
    
    'load data into array
    Open App.Path & "\gas.txt" For Input As #1
    ctr = 0
    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, car(ctr), mileage(ctr)
    Loop
    
    'use an exhaustive search to find total number of cars
    For I = 1 To ctr
        If gasMileage <= mileage(I) And gasMileage > 0 Then
            found = True
            whoctr = whoctr + 1
            picResults.Print car(I); Tab(45); mileage(I)
        End If
    Next I
    
    
    
    'alerts user if no vehicles are found
    If (Not found) Then
        MsgBox ("No cars listed in the array get at least" & " " & gasMileage & " " & "miles per gallon")
    Else
        'gives user final count of vehicles that match or better the mileage given by user
        MsgBox ("There are" & " " & whoctr & " " & "cars listed in the array that get at least" & " " & gasMileage & " " & "miles per gallon")
    End If
    
    Close #1
    
End Sub


Private Sub cmdQuit_Click()
MsgBox ("Thanks for using the Wicked Fun Car Builder, " & " " & UserName & "!")
End
End Sub

