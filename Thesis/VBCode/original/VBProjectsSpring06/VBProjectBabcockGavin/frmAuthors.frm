VERSION 5.00
Begin VB.Form frmAuthors 
   BackColor       =   &H00400000&
   Caption         =   "About Sam and Steve"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   7560
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Go Back to Results Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Image ImgHomies 
      Height          =   2820
      Left            =   2760
      Picture         =   "frmAuthors.frx":0000
      Top             =   1200
      Width           =   3750
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAuthors.frx":464F
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   8775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "About Sam && Steve"
      BeginProperty Font 
         Name            =   "JazzTextExtended"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "frmAuthors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmResults.Show
    frmAuthors.Hide
End Sub
