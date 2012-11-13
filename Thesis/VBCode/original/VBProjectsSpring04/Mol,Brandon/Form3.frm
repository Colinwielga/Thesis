VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF0000&
   Caption         =   "Softails"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form3"
   ScaleHeight     =   8700
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Click to go back to the main menu"
      Height          =   1455
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Picture6 
      Height          =   2055
      Left            =   480
      OLEDropMode     =   1  'Manual
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   2760
      Width           =   3255
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   15
         Left            =   0
         TabIndex        =   11
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   0
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   15
         Left            =   0
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   2055
      Left            =   5880
      Picture         =   "Form3.frx":6065
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   5280
      Width           =   3255
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   480
      Picture         =   "Form3.frx":BF85
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   5280
      Width           =   3255
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   5880
      Picture         =   "Form3.frx":120E3
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   5880
      Picture         =   "Form3.frx":1780C
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   240
      Width           =   3135
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   15
         Left            =   0
         TabIndex        =   7
         Top             =   2040
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   480
      Picture         =   "Form3.frx":1D140
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label10 
      Caption         =   "FXSTS/FXSTSI Springer Softail"
      Height          =   495
      Left            =   5880
      TabIndex        =   15
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label Label9 
      Caption         =   "FXSTB/FXSTBI Night Train"
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label Label8 
      Caption         =   "FXST/FXSTI Softail Standard"
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "FXSTD/FXSTDI Softail Deuce"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "FLSTF/FLSTFI Fat Boy"
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "FLSTC/FLSTCI Heritage Softail Classic"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form3.Hide
End Sub
