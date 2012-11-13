VERSION 5.00
Begin VB.Form fmrWomen 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Questions"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1335
      Left            =   4800
      TabIndex        =   4
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1335
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1335
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblWomen 
      BackColor       =   &H0080C0FF&
      Caption         =   "Welcome to the Women's Section"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   5775
   End
End
Attribute VB_Name = "fmrWomen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
