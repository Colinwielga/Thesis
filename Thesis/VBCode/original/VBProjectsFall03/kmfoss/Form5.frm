VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF8080&
   Caption         =   "Form5"
   ClientHeight    =   8490
   ClientLeft      =   510
   ClientTop       =   450
   ClientWidth     =   12135
   LinkTopic       =   "Form5"
   ScaleHeight     =   8490
   ScaleWidth      =   12135
   Begin VB.CommandButton cmdForm2 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Friday and Saturday: 8pm, 10pm, Midnight, 2am"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   6120
      Width           =   7335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Rounds:  Sunday thru Thursday: 8pm, 10pm, Midnight"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   5520
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "Duty:  Sunday thru Saturday evenings: 8:00pm -8:00am"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   4920
      Width           =   7455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form5.frx":0000
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   11895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form5.frx":0156
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   11775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Duty"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForm2_Click()
Form5.Hide
Form2.Show
End Sub
