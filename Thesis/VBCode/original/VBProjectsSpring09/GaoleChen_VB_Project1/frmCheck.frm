VERSION 5.00
Begin VB.Form frmCheck 
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   Picture         =   "frmCheck.frx":0000
   ScaleHeight     =   6915
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main"
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total price"
      Height          =   615
      Left            =   7320
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      Height          =   3015
      Left            =   6480
      ScaleHeight     =   2955
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "How much is this meal?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Digital Menu
'Form Name: frmCheck
'Authors: Gaole Chen
'Date Written: 3/9/09
'Objective: The user can check out by clicking the button.

Option Explicit

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
frmCheck.Hide
frmWelcome.Show
End Sub

Private Sub cmdTotal_Click()
'the variables are dimmed in module and the total price has been summed up already, so here we can just display.
Dim Tax As Single, Taxrate As Single
Taxrate = 0.08
Tax = Totalcost * Taxrate
picResults.Print "You ordered:", FormatCurrency(FormatNumber(Totalcost), 2)
picResults.Print "Taxes:", FormatCurrency(FormatNumber(Tax), 2)
Totalcost = Totalcost + Tax
picResults.Print "Total cost:", FormatCurrency(FormatNumber(Totalcost), 2)
End Sub
