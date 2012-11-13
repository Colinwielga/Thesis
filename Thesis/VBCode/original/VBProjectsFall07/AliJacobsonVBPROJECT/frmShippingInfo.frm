VERSION 5.00
Begin VB.Form frmShippingInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FF80&
      Caption         =   "Enter Shipping Information then Click Here to Finalize Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   4575
   End
   Begin VB.TextBox txtPhone 
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Top             =   4680
      Width           =   4935
   End
   Begin VB.TextBox txtAddress 
      Height          =   1215
      Left            =   2280
      TabIndex        =   4
      Top             =   3240
      Width           =   4935
   End
   Begin VB.TextBox txtBillingName 
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   2205
      Left            =   960
      Picture         =   "frmShippingInfo.frx":0000
      Top             =   120
      Width           =   2250
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblBillingName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblShipping 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shipping Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4080
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
End
Attribute VB_Name = "frmShippingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdQuit_Click()
Dim BillingName As String

BillingName = txtBillingName.Text

MsgBox "Thank for you for your purchase " & BillingName & " The total cost was : " & FormatCurrency(Total) & " Please shop with us again soon, and have a great day!"

End Sub
