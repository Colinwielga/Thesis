VERSION 5.00
Begin VB.Form frm14 
   Caption         =   "frm14"
   ClientHeight    =   11505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   ScaleHeight     =   11505
   ScaleWidth      =   14955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd60 
      Caption         =   "Total"
      Height          =   975
      Left            =   7320
      TabIndex        =   18
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton cmd56 
      Caption         =   "high to low"
      Height          =   855
      Left            =   11400
      TabIndex        =   17
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton cmd55 
      Caption         =   "Clear"
      Height          =   855
      Left            =   13080
      TabIndex        =   16
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton cmd54 
      Caption         =   "See the price"
      Height          =   855
      Left            =   9600
      TabIndex        =   15
      Top             =   9720
      Width           =   1455
   End
   Begin VB.PictureBox picResults3 
      Height          =   9135
      Left            =   9600
      ScaleHeight     =   9075
      ScaleWidth      =   5115
      TabIndex        =   14
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton cmd53 
      Caption         =   "Leave"
      Height          =   1095
      Left            =   2880
      TabIndex        =   13
      Top             =   10200
      Width           =   1695
   End
   Begin VB.CommandButton cmd52 
      Caption         =   "clear"
      Height          =   1095
      Left            =   480
      TabIndex        =   12
      Top             =   10200
      Width           =   1695
   End
   Begin VB.CommandButton cmd50 
      Caption         =   "Banana smoothie"
      Height          =   975
      Left            =   5160
      TabIndex        =   11
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmd49 
      Caption         =   "Mango smoothie"
      Height          =   975
      Left            =   2880
      TabIndex        =   10
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton cmd48 
      Caption         =   "Iced Tea"
      Height          =   975
      Left            =   480
      TabIndex        =   9
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton cmd47 
      Caption         =   "Iced Mocha"
      Height          =   1095
      Left            =   7320
      TabIndex        =   8
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmd46 
      Caption         =   "Iced Latte"
      Height          =   1095
      Left            =   5160
      TabIndex        =   7
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmd45 
      Caption         =   "Espresso Mocha"
      Height          =   1095
      Left            =   2880
      TabIndex        =   6
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmd44 
      Caption         =   "Espresso"
      Height          =   1095
      Left            =   480
      TabIndex        =   5
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmd43 
      Caption         =   "Mocha"
      Height          =   1095
      Left            =   7320
      TabIndex        =   4
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmd42 
      Caption         =   "Cappuccino"
      Height          =   1095
      Left            =   5160
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmd41 
      Caption         =   "Black Coffee"
      Height          =   1095
      Left            =   2880
      TabIndex        =   2
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmd40 
      Caption         =   "Latte Coffee"
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox picResults2 
      Height          =   5895
      Left            =   360
      ScaleHeight     =   5835
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "frm14.frx":0000
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frm14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Single


Private Sub Command4_Click()

End Sub

Private Sub Command8_Click()

End Sub

Private Sub cmd40_Click()
picResults2.Print "Latte Coffee             $3.75"
total = total + 3.75
End Sub

Private Sub cmd41_Click()
picResults2.Print "Black Coffee            $4.15"
total = total + 4.15
End Sub

Private Sub cmd42_Click()
picResults2.Print "Cappucino               $5.99"
total = total + 5.99
End Sub

Private Sub cmd43_Click()
picResults2.Print "Mocha                    $4.99"
total = total + 4.99
End Sub

Private Sub cmd44_Click()
picResults2.Print "Espresso                 $4.99"
total = total + 4.99
End Sub

Private Sub cmd45_Click()
picResults2.Print "Espresso Mocha       $6.99"
total = total + 6.99
End Sub

Private Sub cmd46_Click()
total = total + 4.15
picResults2.Print "Iced Latte               $4.15"
End Sub

Private Sub cmd47_Click()
total = total + 5.15
picResults2.Print "Iced Mocha              $5.15"
End Sub

Private Sub cmd48_Click()
total = total + 4.99
picResults2.Print "Iced Tea                $4.99"
End Sub

Private Sub cmd49_Click()
total = total + 5.99
picResults2.Print "Mango snoothie      $5.99"
End Sub

Private Sub cmd50_Click()
total = total + 5.99
picResults2.Print "Banana smoothie     $5.99"
End Sub

Private Sub cmd52_Click()
picResults2.Cls
total = 0
End Sub

Private Sub cmd53_Click()
frm14.Visible = False
frm15.Visible = True
End Sub

Private Sub cmd54_Click()
Dim CTR As Single
Dim Sname(1 To 100) As String
Dim Sprice(1 To 100) As Double
Open App.Path & "\Coffee Price.txt" For Input As #1
CTR = 0
Do Until EOF(1)
CTR = CTR + 1
Input #1, Sname(CTR), Sprice(CTR)
picResults3.Print Sname(CTR), FormatCurrency(Sprice(CTR))
Loop
Close #1


End Sub

Private Sub cmd55_Click()
picResults3.Cls
End Sub

Private Sub cmd56_Click()
Dim pos As Single
Dim Arr(1 To 100) As Double
Dim CTR As Single
Dim pass As Double












Dim temp As Single
Dim Sname(1 To 100) As String
Dim Sprice(1 To 100) As Double
Open App.Path & "\Coffee Price.txt" For Input As #1
CTR = 0
Do Until EOF(1)
CTR = CTR + 1
Input #1, Sname(CTR), Sprice(CTR)
Loop
Close #1
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If Arr(pos) > Arr(pos + 1) Then
        temp = Arr(pos)
        Arr(pos) = Arr(pos + 1)
        Arr(pos + 1) = temp
        End If
    Next pos
Next pass
For pos = 1 To CTR
picResults3.Print Sname(pos), FormatCurrency(Sprice(pos))
Next pos
End Sub

Private Sub cmd60_Click()
picResults2.Print "Total                         " & FormatCurrency(total)
End Sub
