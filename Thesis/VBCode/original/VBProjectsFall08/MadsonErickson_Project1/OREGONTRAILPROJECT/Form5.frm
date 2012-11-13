VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "RETURN TO MAIN PAGE"
      Height          =   1815
      Left            =   4320
      TabIndex        =   1
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "We read a lot about Oregon Trail on Wikipedia"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   12135
   End
   Begin VB.Label Label3 
      Caption         =   $"Form5.frx":0000
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   12135
   End
   Begin VB.Label Label2 
      Caption         =   $"Form5.frx":009D
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   12135
   End
   Begin VB.Label Label1 
      Caption         =   "WORKS CITED"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Hide
Form2.Show

End Sub
