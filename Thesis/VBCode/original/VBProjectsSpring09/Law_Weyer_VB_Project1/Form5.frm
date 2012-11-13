VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   LinkTopic       =   "Form5"
   ScaleHeight     =   3090
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Exchange"
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtCurrency 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblDesire 
      BackStyle       =   0  'Transparent
      Caption         =   "Desired Currency:"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblUS 
      BackStyle       =   0  'Transparent
      Caption         =   "U.S. amount to be exchanged:"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblEuro 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Euro"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblForint 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Hungarian Forint"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblCzech 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Czech Coruna"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Form5.Hide
End Sub

Private Sub cmdConvert_Click()
Dim amount As Single, currencies(1 To 50) As String, money(1 To 50) As Single
Dim K As Integer, Ctr As Integer, i As Integer

Dim answer As Single
Ctr = 0

amount = txtAmount.Text
K = txtCurrency.Text

Open App.Path & "\currency.txt" For Input As #1

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, currencies(Ctr), money(Ctr)
Loop
Close #1

answer = amount * money(txtCurrency.Text)
MsgBox " You can exchange " & FormatCurrency(amount, 0) & "  U.S. Dollar for " & FormatCurrency(answer, 0) & " " & currencies(K), , "Exchange"

End Sub

