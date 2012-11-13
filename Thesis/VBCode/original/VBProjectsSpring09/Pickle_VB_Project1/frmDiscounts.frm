VERSION 5.00
Begin VB.Form frmAccesories 
   Caption         =   "Form1"
   ClientHeight    =   10605
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate My Accessories"
      Height          =   1335
      Left            =   3240
      TabIndex        =   9
      Top             =   4440
      Width           =   3015
   End
   Begin VB.PictureBox picResultsTwo 
      Height          =   1215
      Left            =   6960
      ScaleHeight     =   1155
      ScaleWidth      =   5115
      TabIndex        =   8
      Top             =   5640
      Width           =   5175
   End
   Begin VB.TextBox txtThree 
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtTwo 
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtOne 
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Final Screen"
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1335
      Left            =   3360
      TabIndex        =   2
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read The File"
      Height          =   1335
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.PictureBox picResults 
      Height          =   3135
      Left            =   6960
      ScaleHeight     =   3075
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label lblDiscounts 
      Caption         =   "Enter Desiered Accessories.  (Limit 3 Per Person)"
      Height          =   735
      Left            =   3000
      TabIndex        =   7
      Top             =   3360
      Width           =   3615
   End
End
Attribute VB_Name = "frmAccesories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
Dim Subtotal As Double
Subtotal = 0
    Select Case Calculate
        Case 1
            Subtotal = Subtotal + 9.99
            runningtotal = runningtotal + Subtotal
        Case 2
            Subtotal = Subtotal + 25.99
            runningtotal = runningtotal + Subtotal
        Case 3
            Subtotal = Subtotal + 15.95
            runningtotal = runningtotal + Subtotal
        Case 4
            Subtotal = Subtotal + 9.98
            runningtotal = runningtotal + Subtotal
        Case 5
            Subtotal = Subtotal + 69.99
            runningtotal = runningtotal + Subtotal
        Case 6
            Subtotal = Subtotal + 5.99
            runningtotal = runningtotal + Subtotal
        Case 7
            Subtotal = Subtotal + 149.99
            runningtotal = runningtotal + Subtotal
        Case 8
            Subtotal = Subtotal + 19.95
            runningtotal = runningtotal + Subtotal
        Case 9
            Subtotal = Subtotal + 7.99
            runningtotal = runningtotal + Subtotal
        Case 10
            Subtotal = Subtotal + 45.99
            runningtotal = runningtotal + Subtotal
        End Select
    picResultsTwo.Print "You subtotal thus far, including accessories is about"; FormatCurrency(runningtotal)
    
        
            
        
End Sub

Private Sub cmdNext_Click()
    frmDiscounts.Hide
    frmReceipt.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRead_Click()
Dim NameAccessory(1 To 10) As String
Dim Price(1 To 10) As Double
Dim Counter As Double
Counter = 0
picResults.Print "Number", "Accessory", "Price"
picResults.Print "************************************"
Open App.Path & "\Accessories.txt" For Input As #2
        Do Until EOF(2)
            Counter = Counter + 1
            Input #2, NameAccessory(Counter), Price(Counter)
            picResults.Print Counter, NameAccessory(Counter), FormatCurrency(Price(Counter))
        Loop
End Sub

Private Sub txtOne_Change()
    One = txtOne.Text
End Sub
