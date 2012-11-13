VERSION 5.00
Begin VB.Form frmReceive 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   3825
   ClientTop       =   3900
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   8370
   Begin VB.CommandButton cmdMakeReception 
      Caption         =   "Make Transaction"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lbl5 
      Height          =   495
      Left            =   7080
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl4 
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl3 
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl2 
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl1 
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Enter Number of Place to Receive Money From:"
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "= 5"
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "= 4"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "= 3"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "= 2"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "= 1"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Who are you taking money from?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "frmReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form decides who to take the money from
'It first adds the money to the current player's account, then it subtracts if from the correct place
Private Sub cmdMakeReception_Click()
Dim Amount As Single
Amount = frmBoard.txtTransaction(Turn - 1).Text
Number = Text1.Text
PlayerMoney(Turn) = PlayerMoney(Turn) + Amount
frmBoard.picMoney(Turn - 1).Cls
frmBoard.picMoney(Turn - 1).Print "You Have "; FormatCurrency(PlayerMoney(Turn))
'if the money is coming from a player it will adjust the player's account
If Number <= 4 Then
    PlayerMoney(Number) = PlayerMoney(Number) - Amount
    frmBoard.picMoney(Number - 1).Cls
    frmBoard.picMoney(Number - 1).Print "You Have "; FormatCurrency(PlayerMoney(Number))
    'Stat(Position(Turn) + 1) = 1
    frmPay.Visible = False
    frmBoard.cmdPlayerPay(Turn - 1).Enabled = False
    frmBoard.txtTransaction(Turn - 1).Enabled = False
    frmBoard.cmdOk.Enabled = True
'the else if for receiving from the bank
Else
    frmPay.Visible = False
    frmBoard.cmdPlayerPay(Turn - 1).Enabled = False
    frmBoard.txtTransaction(Turn - 1).Enabled = False
    frmBoard.cmdOk.Enabled = True
End If
frmReceive.Visible = False
frmBoard.cmdPlayerReceive(Turn - 1).Enabled = False
frmBoard.cmdPlayerPay(Turn - 1).Enabled = False
frmBoard.Visible = True
End Sub



Private Sub Text1_Change()
    If Text1 <> "" Then
        cmdMakeReception.Enabled = True
    End If
    
End Sub
