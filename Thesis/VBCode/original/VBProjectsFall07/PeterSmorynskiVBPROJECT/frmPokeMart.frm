VERSION 5.00
Begin VB.Form frmPokeMart 
   BackColor       =   &H8000000D&
   Caption         =   "POKEMART "
   ClientHeight    =   10845
   ClientLeft      =   510
   ClientTop       =   225
   ClientWidth     =   11970
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   11970
   Begin VB.CommandButton cmdRtnHub 
      Caption         =   "Return to Pokemon Central"
      Height          =   1095
      Left            =   7680
      TabIndex        =   15
      Top             =   9600
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   6735
      Left            =   5520
      ScaleHeight     =   6675
      ScaleWidth      =   6195
      TabIndex        =   7
      Top             =   2520
      Width           =   6255
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Cart and Start Over "
      Height          =   1095
      Left            =   5520
      TabIndex        =   6
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   1095
      Left            =   3240
      TabIndex        =   5
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   9840
      TabIndex        =   4
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton cmdRepel 
      Caption         =   "Repel"
      Height          =   1335
      Left            =   3240
      TabIndex        =   3
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdPotion 
      Caption         =   "Potion"
      Height          =   1455
      Left            =   3240
      TabIndex        =   2
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdUltraball 
      BackColor       =   &H80000015&
      Caption         =   "Ultraball"
      Height          =   1575
      Left            =   3240
      TabIndex        =   1
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdPokeball 
      Caption         =   "Pokeball"
      Height          =   1455
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblcart 
      Caption         =   "PokeMart Shopping Cart"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000013&
      Caption         =   "Repel = $5.10"
      Height          =   495
      Left            =   10680
      TabIndex        =   14
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      Caption         =   "Potion = $7.50"
      Height          =   495
      Left            =   9480
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Caption         =   "Ultraball = $54.99"
      Height          =   495
      Left            =   8040
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "Pokeball = $19.99"
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "INVENTORY:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   $"frmPokeMart.frx":0000
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   9
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H80000013&
      Caption         =   "YOUR ACCOUNT: $200 POKEDOLLARS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   4710
      Left            =   3600
      Picture         =   "frmPokeMart.frx":00C0
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "frmPokeMart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'inspired by Lab7 pt.2 with addition features
Dim Subtotal As Single, CTR As Integer, Account As Single, PokeballCTR As Integer
Private Sub cmdPokeball_Click() 'add pokeball to cart with math function
Dim Pokeball As Single
    Pokeball = 19.99
    picResults.Print "Pokeball", FormatCurrency(Pokeball)
    Subtotal = Subtotal + Pokeball
End Sub

Private Sub cmdUltraball_Click() 'add ultraball to cart with math function
Dim Ultraball As Single
Ultraball = 54.99
picResults.Print "Ultraball", FormatCurrency(Ultraball)
Subtotal = Subtotal + Ultraball

End Sub

Private Sub cmdPotion_Click() 'add potion to cart with math function
Dim Potion As Single
Potion = 7.5
picResults.Print "Potion", FormatCurrency(Potion)
Subtotal = Subtotal + Potion
End Sub

Private Sub cmdRepel_Click() 'add repel to cart with math function
Dim Repel As Single
Repel = 5.1
picResults.Print "Repel", FormatCurrency(Repel)
Subtotal = Subtotal + Repel
End Sub

Private Sub cmdTotal_Click() 'calculate the totals
Dim Total As Single, Tax As Single
Tax = Subtotal * 0.08
Total = Tax + Subtotal
Account = 200 - Total
If Total > 200 Then
   MsgBox ("Oops! You don't have enough Pokedollars! You should START OVER and watch your money better..."), , ("overdrawn!")
   Else
picResults.Print "---------------"
picResults.Print "Subtotal", FormatCurrency(Subtotal)
picResults.Print "Tax", FormatCurrency(Tax)
picResults.Print "Total", FormatCurrency(Total)
picResults.Print "Account", FormatCurrency(Account)
End If
End Sub

Private Sub cmdClear_Click() 'reset form
Subtotal = 0
Account = 200
picResults.Cls
End Sub
Private Sub cmdRtnHub_Click() 'return to Pokemon Central
cmdTotal.Visible = False
cmdClear.Visible = False
cmdPokeball.Visible = False
cmdUltraball.Visible = False
cmdPotion.Visible = False
cmdRepel.Visible = False
frmPokeMart.Hide
frmCentralHub.Show
MsgBox ("Welcome back to Pokemon Central! The PokeMart is essential for Pokemon Trainers, and teaches youngsters a thing or two about money management."), , ("INSTRUCTION: CHOOSE YOUR NEXT DESTINATION!")
End Sub

Private Sub cmdQuit_Click()
End
End Sub


