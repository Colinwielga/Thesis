VERSION 5.00
Begin VB.Form frmNoPay 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd3 
      BackColor       =   &H0080FF80&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdFree 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Staying at my girlfriend's house saved you money! Now Lets get you a sweet rental car."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   4575
   End
End
Attribute VB_Name = "frmNoPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Quit Button'
Private Sub cmdEnd3_Click()
    End
End Sub

'girlfriends house was free, so taking you to the rental car page'
'Life is more fun with options'
'October 14 2009'
'Blake Bauer'
Private Sub cmdFree_Click()
    frmNoPay.Hide
    frmCars.Show
End Sub
