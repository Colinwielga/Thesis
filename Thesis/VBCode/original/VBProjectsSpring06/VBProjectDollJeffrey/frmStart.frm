VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00FF0000&
   Caption         =   "Decathalon Scoring Converter"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3240
   FillColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1728.396
   ScaleMode       =   0  'User
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H000000FF&
      Caption         =   "Click HERE to begin"
      Height          =   495
      Left            =   360
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblobj 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmStart.frx":0000
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblBy 
      BackColor       =   &H00FF0000&
      Caption         =   "Created by Jeff Doll"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   360
      Picture         =   "frmStart.frx":00CA
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF0000&
      Caption         =   "Decathlon Scoring        Converter"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdStart_Click()
    'The name will be stored in n and will appear throught the project
    'Also the first form will be hidden and the next form will appear
    n = InputBox("Enter your name", "Enter your name", "Enter your name")
    frmStart.Hide
    frmInput.Show
End Sub
