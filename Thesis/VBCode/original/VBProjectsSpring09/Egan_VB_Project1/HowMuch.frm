VERSION 5.00
Begin VB.Form frmHowMuch 
   BackColor       =   &H0000C000&
   Caption         =   "Welcome!"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoToHomepage 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go To Homepage"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "Get Credit"
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtMoney 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblTotalAmount 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblHowMuch 
      BackColor       =   &H0000C000&
      Caption         =   "How much money would you like to start with: $1, $2, $5, or $10?"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmHowMuch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Sports Betting Project
    'frmHowmuch
    'Written by: Sean Egan
    'Written on: 3/22/09
    'This form prompts the user to choose how much money
    ' they would like to start with.
    
Private Sub cmdExit_Click()
    'This ends the program
    End
End Sub

Private Sub cmdGetMoney_Click()
    'Sets what is entered into the text box equal to the
    ' global variable "Total"
    Total = txtMoney.Text
    
    'If/ElseIf statement that gives the user $1, $2, $5, or $10,
    ' depending on what they choose. If they do not enter one of
    ' these numbers, Message Box will appear telling them to try
    ' again. They will not be able to advance in the program unless
    ' they choose a value. Also, they are only allowed to choose
    ' one value.
    If Total = 1 Then
            lblTotalAmount.Caption = FormatCurrency(Total)
            cmdGoToHomepage.Enabled = True
            cmdGetMoney.Enabled = False
        ElseIf Total = 2 Then
            lblTotalAmount.Caption = FormatCurrency(Total)
            cmdGoToHomepage.Enabled = True
            cmdGetMoney.Enabled = False
        ElseIf Total = 5 Then
            lblTotalAmount.Caption = FormatCurrency(Total)
            cmdGoToHomepage.Enabled = True
            cmdGetMoney.Enabled = False
        ElseIf Total = 10 Then
            lblTotalAmount.Caption = FormatCurrency(Total)
            cmdGoToHomepage.Enabled = True
            cmdGetMoney.Enabled = False
        Else
            MsgBox ("I'm sorry. Only enter a 1, 2, 5, or 10 in the text box.")
            Total = 0
    End If
End Sub

Private Sub cmdGoToHomepage_Click()
    'Hides the current form
    frmHowMuch.Hide
    'Loads the Homepage
    frmHomepage.Show
End Sub
