VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FF8080&
   Caption         =   "Form4"
   ClientHeight    =   8715
   ClientLeft      =   285
   ClientTop       =   660
   ClientWidth     =   12195
   LinkTopic       =   "Form4"
   ScaleHeight     =   8715
   ScaleWidth      =   12195
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   4
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdActive 
      Caption         =   "Active Programming"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   3
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton cmdPassive 
      Caption         =   "Passive Programming"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   2
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "There are two types of Programming, Passive and Active"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   4200
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form4.frx":0000
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActive_Click()
Form4.Hide
Form11.Show
End Sub

Private Sub cmdMenu_Click()
Form4.Hide
Form2.Show
End Sub

Private Sub cmdPassive_Click()
Form4.Hide
Form10.Show
End Sub
