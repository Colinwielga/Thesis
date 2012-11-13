VERSION 5.00
Begin VB.Form frmPay 
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   2850
   ClientTop       =   3570
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   10170
   Begin VB.CommandButton cmdMakeTransaction 
      Caption         =   "Make Transaction"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Enter Number  of Place of Payment:"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "= 5"
      Height          =   495
      Left            =   9000
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "= 4"
      Height          =   495
      Left            =   7080
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "= 3 "
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "= 2"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "= 1"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblB 
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblP4 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblP3 
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblP2 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblP1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Who Gets Paid?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   7695
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Amount As Single
Private Sub cmdMakeTransaction_Click()
 Amount = frmBoard.txtTransaction(Turn - 1).Text
 Number = Text1.Text
PlayerMoney(Turn) = PlayerMoney(Turn) - Amount
frmBoard.picMoney(Turn - 1).Cls
frmBoard.picMoney(Turn - 1).Print "You Have "; (PlayerMoney(Turn))
If Number <= 4 Then
    PlayerMoney(Number) = PlayerMoney(Number) + Amount
    frmBoard.picMoney(Number - 1).Cls
    frmBoard.picMoney(Number - 1).Print "You Have "; (PlayerMoney(Number))
    frmBoard.picPlayer(Turn - 1).Print Place(Position(Turn) + 1)
    Stat(Position(Turn) + 1) = 1
    frmPay.Visible = False
    frmBoard.cmdPlayerPay(Turn - 1).Enabled = False
    
    frmBoard.txtTransaction(Turn - 1).Enabled = False
    frmBoard.cmdOk.Enabled = True
    
Else
    frmBoard.picPlayer(Turn - 1).Print Place(Position(Turn) + 1)
    Stat(Position(Turn) + 1) = 1
    frmPay.Visible = False
    frmBoard.cmdPlayerPay(Turn - 1).Enabled = False
    
    frmBoard.txtTransaction(Turn - 1).Enabled = False
    frmBoard.cmdOk.Enabled = True
End If
End Sub

Private Sub lblP3_Click()
    lblP3.Caption = Player(3)
End Sub

Private Sub lblP4_Click()
    lblP4.Caption = Player(4)
End Sub
    
Private Sub Text1_Change()
    If Text1 <> "" Then
        cmdMakeTransaction.Enabled = True
    End If
    
End Sub
