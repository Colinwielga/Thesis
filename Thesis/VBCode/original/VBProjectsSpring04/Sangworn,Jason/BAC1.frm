VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF0000&
   Caption         =   "Form5"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9885
   LinkTopic       =   "Form5"
   ScaleHeight     =   9435
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here to Begin"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   2
      Top             =   7560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "BAC1.frx":0000
      Top             =   2040
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Blood Alcohol Content"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   900
      Left            =   1830
      TabIndex        =   1
      Top             =   720
      Width           =   6825
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Command1_Click()
Form5.Hide
Form1.Show

End Sub
