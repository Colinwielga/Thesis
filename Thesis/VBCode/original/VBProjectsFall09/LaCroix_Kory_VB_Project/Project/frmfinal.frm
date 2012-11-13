VERSION 5.00
Begin VB.Form frmfinal 
   BackColor       =   &H00400040&
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   15510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   2895
      Left            =   9120
      ScaleHeight     =   2835
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   3360
      Width           =   5775
   End
   Begin VB.CommandButton cmdfinaltotal 
      Caption         =   "Final Total"
      Height          =   735
      Left            =   11040
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   855
      Left            =   11160
      TabIndex        =   2
      Top             =   6600
      Width           =   1935
   End
   Begin VB.PictureBox picbrett1 
      Height          =   5295
      Left            =   360
      Picture         =   "frmfinal.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   2280
      Width           =   7935
   End
   Begin VB.Label lblfinal 
      BackColor       =   &H00400040&
      Caption         =   "You are now set to be a member of the Brett Favre Fan Club. Awesome. Go Vikes!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "frmfinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Brett Favre Fan Club
'Form Name: frmfinal
'Author: Kory LaCroix
'Date Written: 10/19/08
'Objective: To end the program and give the final total cost of the trip
Option Explicit
Private Sub cmdEnd_Click()
'this ends the program
End
End Sub

Private Sub cmdfinaltotal_Click()
'this is a brief message to the user
picResults.Print "The price you will pay is well worth it."
picResults.Print "Maybe this is the year the Vikes will win it all."
picResults.Print "  "
'this message will display the total cost of the trip which has been adding up throughout the program
picResults.Print "The total cost of your gear. hotel, tickets, and flight will be "; FormatCurrency(runningtotal)
End Sub

