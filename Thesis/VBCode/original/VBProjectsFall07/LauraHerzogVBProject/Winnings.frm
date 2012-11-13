VERSION 5.00
Begin VB.Form Winnings 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   Picture         =   "Winnings.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdwinnings 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Click Here to see everything you have won!!!!!"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      ScaleHeight     =   1635
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   3960
      Width           =   5535
   End
End
Attribute VB_Name = "Winnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdwinnings_Click()
picResults.Print (WholeName) & " You have won " & FormatCurrency(Runningtotal)
If TV = True Then
    picResults.Print "AND A NEW TELEVISION"
End If
If Runningtotal = 0 Then
    picResults.Print "You have won Zero Dollars"
End If
End Sub
