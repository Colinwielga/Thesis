VERSION 5.00
Begin VB.Form frmForm1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   5175
   ClientTop       =   5685
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6030
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Begin!"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblName 
      Caption         =   "Please Enter Your Name Here"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBegin_Click()
MsgBox "Hello " & txtName.Text & "!  Welcome to the Virtual Match-Maker 2.0!"
frmForm1.Hide
frmForm2.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
