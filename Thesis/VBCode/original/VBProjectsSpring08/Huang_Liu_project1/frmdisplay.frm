VERSION 5.00
Begin VB.Form frmdisplay 
   Caption         =   "Message Board"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form2"
   Picture         =   "frmdisplay.frx":0000
   ScaleHeight     =   8700
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   360
      ScaleHeight     =   4035
      ScaleWidth      =   8595
      TabIndex        =   5
      Top             =   3360
      Width           =   8655
   End
   Begin VB.CommandButton cmdsummit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Summit Your Response!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7200
      Picture         =   "frmdisplay.frx":4B80B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtresponse 
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   6855
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to the main page"
      Height          =   735
      Left            =   7320
      TabIndex        =   0
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label lbl2 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Message Board"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   8655
   End
   Begin VB.Label lbl1 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Type your message here!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   6855
   End
End
Attribute VB_Name = "frmdisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Response As String


Private Sub cmdback_Click()
frmdisplay.Visible = False
frmmain.Visible = True
End Sub

Private Sub cmdsummit_Click()
Response = txtresponse.Text
picoutput.Print "______________________________________________________________________________"
picoutput.Print Response
picoutput.Print
picoutput.Print "Message written by" & Customer & " on " & Dates
picoutput.Print
End Sub

