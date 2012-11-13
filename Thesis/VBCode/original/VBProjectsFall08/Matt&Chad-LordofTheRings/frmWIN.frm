VERSION 5.00
Begin VB.Form frmWIN 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   2220
   ClientTop       =   1710
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   Picture         =   "frmWIN.frx":0000
   ScaleHeight     =   6645
   ScaleWidth      =   9600
   Begin VB.CommandButton Command1 
      Caption         =   "The Journey Has Ended"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWIN.frx":EC03
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   9255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Congragulations!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub
