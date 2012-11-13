VERSION 5.00
Begin VB.Form Tower 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Tower"
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
      Height          =   615
      Left            =   17280
      TabIndex        =   16
      Top             =   13080
      Width           =   1215
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14880
      TabIndex        =   15
      Top             =   13080
      Width           =   2055
   End
   Begin VB.PictureBox Picture7 
      Height          =   2415
      Left            =   6120
      Picture         =   "Tower.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   13
      Top             =   10680
      Width           =   2775
   End
   Begin VB.PictureBox Picture8 
      Height          =   1815
      Left            =   12960
      Picture         =   "Tower.frx":505A
      ScaleHeight     =   1755
      ScaleWidth      =   3555
      TabIndex        =   11
      Top             =   10920
      Width           =   3615
   End
   Begin VB.PictureBox Picture6 
      Height          =   3255
      Left            =   14160
      Picture         =   "Tower.frx":6BAE
      ScaleHeight     =   3195
      ScaleWidth      =   4635
      TabIndex        =   10
      Top             =   4800
      Width           =   4695
   End
   Begin VB.PictureBox Picture5 
      Height          =   2655
      Left            =   15360
      Picture         =   "Tower.frx":D3CA
      ScaleHeight     =   2595
      ScaleWidth      =   2355
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.PictureBox Picture4 
      Height          =   6015
      Left            =   6240
      Picture         =   "Tower.frx":127EE
      ScaleHeight     =   5955
      ScaleWidth      =   7515
      TabIndex        =   8
      Top             =   3960
      Width           =   7575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1695
      Left            =   8520
      Picture         =   "Tower.frx":1AC9F
      ScaleHeight     =   1635
      ScaleWidth      =   5235
      TabIndex        =   6
      Top             =   2040
      Width           =   5295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   2640
      Picture         =   "Tower.frx":1E021
      ScaleHeight     =   1995
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   3000
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   9495
      Left            =   480
      Picture         =   "Tower.frx":252A1
      ScaleHeight     =   9435
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   13800
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"Tower.frx":29847
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   9240
      TabIndex        =   14
      Top             =   10680
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "The Tower Bridge at night, while driving across it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15000
      TabIndex        =   12
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "This is what may people refer to as the London Bridge at night.    This is actually named the Tower Bridge, today."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   7
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Famous Sites of the Tower District in London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Tower.frx":299B1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   2880
      TabIndex        =   4
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inside the Monument"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "The Monument"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "Tower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: Tower (Tower.frm)
'Author:Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: The purpose of this form is for the user to learn the history of The Monument, The Tower Bridge(London
                'Bridge), and the Tower of London
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'User returns back to choose a new district on the Map of London page
Tower.Hide
MapLondon.Show
End Sub
