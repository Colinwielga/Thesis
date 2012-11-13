VERSION 5.00
Begin VB.Form frmPay 
   Caption         =   "Gross Income"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   1575
      Left            =   360
      ScaleHeight     =   1515
      ScaleWidth      =   4515
      TabIndex        =   6
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txtPayrate 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtHours 
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblPayrate 
      Caption         =   "Insert Pay Rate"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblHours 
      Caption         =   "Insert Hours Worked"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGo_Click()
Dim Pay As Integer
Dim Hours As Single
Dim Payrate As Single

Hours = txtHours.Text
Payrate = txtPayrate.Text

    If Hours > 40 Then
    Pay = Hours * (Payrate * 1.5)
    picResults.Print "You have earned-"; FormatCurrency(Pay); "for"; (Hours); " hours of work at a base pay rate of"; FormatCurrency(Payrate); "per hour"
    
    ElseIf Hours <= 40 Then
    Pay = Hours * Payrate
    picResults.Print "You have earned-"; FormatCurrency(Pay); "for"; (Hours); "of work at a base pay rate of"; FormatCurrency(Payrate); "per hour"
    
    
    End If
End Sub
