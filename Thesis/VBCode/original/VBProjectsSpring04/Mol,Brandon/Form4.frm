VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H0000FFFF&
   Caption         =   "Sportsters"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form4"
   ScaleHeight     =   8640
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Click to go back to the main menu"
      Height          =   1335
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   5880
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   3
      Top             =   4440
      Width           =   3135
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   600
      Picture         =   "Form4.frx":574E
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   4440
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   5880
      Picture         =   "Form4.frx":A985
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   600
      Picture         =   "Form4.frx":FCE5
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "XL 1200R Sportster 1200 Roadster"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "XL 1200C Sportster 1200 Custom"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "XL 883C Sportster 883 Custom"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "XL 883 Sportster 883"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   3135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form4.Hide
End Sub
