VERSION 5.00
Begin VB.Form frmLearn 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   Picture         =   "frmLearn.frx":0000
   ScaleHeight     =   9015
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to Home Screen"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4560
      Picture         =   "frmLearn.frx":3628BA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdWR 
      BackColor       =   &H0000C000&
      Caption         =   "Wide Reciever"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4800
      Picture         =   "frmLearn.frx":36347D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H0000C000&
      Caption         =   "Running Back"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8520
      Picture         =   "frmLearn.frx":3643C6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdFull 
      BackColor       =   &H0000C000&
      Caption         =   "Fullback"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8640
      Picture         =   "frmLearn.frx":36576E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdQB 
      BackColor       =   &H0000C000&
      Caption         =   "Quarterback"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1080
      Picture         =   "frmLearn.frx":3660E4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdOline 
      BackColor       =   &H0000C000&
      Caption         =   "The O-Line"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   960
      Picture         =   "frmLearn.frx":366C97
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblLearn 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Learn about the Offensive Positions by clicking one of the pictures!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   6495
   End
End
Attribute VB_Name = "frmLearn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFull_Click()
    frmLearn.Hide
    frmFull.Show
End Sub

Private Sub cmdHome_Click()
    frmLearn.Hide
    frmRoster.Show
    
End Sub

Private Sub cmdOline_Click()
    frmLearn.Hide
    frmOLine.Show
End Sub

Private Sub cmdQB_Click()
    frmLearn.Hide
    frmQB.Show
End Sub

Private Sub cmdRun_Click()
    frmLearn.Hide
    frmRB.Show
End Sub

Private Sub cmdWR_Click()
    frmLearn.Hide
    frmWR.Show
    
End Sub
