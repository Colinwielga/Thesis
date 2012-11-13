VERSION 5.00
Begin VB.Form FrmIndex 
   BackColor       =   &H80000012&
   Caption         =   "Form Index"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture10 
      Height          =   1455
      Left            =   6720
      Picture         =   "FrmIndex.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Picture9 
      Height          =   1455
      Left            =   3720
      Picture         =   "FrmIndex.frx":54A2
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.PictureBox Picture8 
      Height          =   1455
      Left            =   5160
      Picture         =   "FrmIndex.frx":A944
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.PictureBox Picture7 
      Height          =   1455
      Left            =   5160
      Picture         =   "FrmIndex.frx":FDE6
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Picture6 
      Height          =   1455
      Left            =   3720
      Picture         =   "FrmIndex.frx":15288
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      Height          =   1455
      Left            =   6720
      Picture         =   "FrmIndex.frx":1A72A
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   6720
      Picture         =   "FrmIndex.frx":1FBCC
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   3720
      Picture         =   "FrmIndex.frx":2506E
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Index           =   0
      Left            =   5160
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
      Begin VB.PictureBox Picture2 
         Height          =   5055
         Index           =   1
         Left            =   0
         Picture         =   "FrmIndex.frx":2A510
         ScaleHeight     =   4995
         ScaleWidth      =   6315
         TabIndex        =   5
         Top             =   0
         Width           =   6375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   3960
      Picture         =   "FrmIndex.frx":2F8C2
      ScaleHeight     =   675
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Click to Return to Previous Page"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "The Defence"
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
      Left            =   1800
      TabIndex        =   16
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "The Offensive Line"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "The Big 3"
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
      Left            =   2280
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "The Minnesota Vikings!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Creator: Adam Hanson"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   9000
      Width           =   1815
   End
End
Attribute VB_Name = "FrmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Index (FrmIndex)
Private Sub cmdreturn_Click()
    FrmIndex.Hide
    FrmNFL.Show
End Sub
