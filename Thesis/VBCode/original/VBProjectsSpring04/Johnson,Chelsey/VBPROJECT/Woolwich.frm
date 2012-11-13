VERSION 5.00
Begin VB.Form Woolwich 
   BackColor       =   &H00808000&
   Caption         =   "Woolwich"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   17040
      TabIndex        =   7
      Top             =   12840
      Width           =   1695
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13800
      TabIndex        =   6
      Top             =   12840
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   1080
      Picture         =   "Woolwich.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   17115
      TabIndex        =   2
      Top             =   7320
      Width           =   17175
   End
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   6360
      Picture         =   "Woolwich.frx":8EE8
      ScaleHeight     =   5115
      ScaleWidth      =   7635
      TabIndex        =   1
      Top             =   1440
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Text            =   "The most predominent site in Woolwich District is that of the Thmaes Barrier."
      Top             =   240
      Width           =   7575
   End
   Begin VB.Label Label4 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   14280
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Woolwich.frx":1551C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   14760
      TabIndex        =   5
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Thames Barrier has been described as the eighth wonder of the world."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   10440
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Woolwich.frx":1581E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
End
Attribute VB_Name = "Woolwich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: Woolwich (Woolwich.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: The purpose of this form was to inform the user of the history of the Thmaes Barrier in the
                'district of Woolwich
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'User returns to the Map of London Page to choose a new district to view
Woolwich.Hide
MapLondon.Show
End Sub
