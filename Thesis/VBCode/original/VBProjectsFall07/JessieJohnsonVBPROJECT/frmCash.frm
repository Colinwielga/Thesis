VERSION 5.00
Begin VB.Form frmCash 
   BackColor       =   &H00400040&
   Caption         =   "The Campground!"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   Picture         =   "frmCash.frx":0000
   ScaleHeight     =   4905
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picTotal 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      ScaleHeight     =   435
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdGrandTotal 
      Caption         =   "Your grand total is:"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   2640
      Width           =   4575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Thank you for shopping at the Campground! Click here to exit the store."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAmountPaid 
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblHowMuch 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "How much are you paying? Enter amount without the dollar sign.   (Ex: 49.50)"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "frmCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGrandTotal_Click()
'reminds the user of their total, just in case they forgot!
picTotal.Print FormatCurrency(GrandTotal, 2)
End Sub

Private Sub cmdPay_Click()
'the user enters the amount they would like to pay, and either get change back, give
'exact change, or are asked to recount and try again
'if they either pay over or pay using exact change, they are able to use the Quit
'button and leave the program using the Visible property
Dim Change As Single
Dim AmountPaid As Single
AmountPaid = txtAmountPaid.Text
If AmountPaid > GrandTotal Then
    Change = AmountPaid - GrandTotal
    picResults.Print FormatCurrency(Change); " is your change."
    cmdQuit.Visible = True
ElseIf AmountPaid = GrandTotal Then
    picResults.Print "Exact change! How nice!"
    cmdQuit.Visible = True
Else
    MsgBox "Count your money and try again."
End If
End Sub

Private Sub cmdQuit_Click()
'gives the user a farewell message before ending the program
MsgBox "Have a nice day!"
End
End Sub
