VERSION 5.00
Begin VB.Form frmFirst 
   Caption         =   "StartPage"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8895
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   8835
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bail Out"
         DisabledPicture =   "Form1.frx":1B409
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6840
         MaskColor       =   &H00FFFFC0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6000
         Width           =   3015
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Start Your Descent"
         DisabledPicture =   "Form1.frx":36812
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6000
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Click For Music"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   6720
         TabIndex        =   6
         Top             =   3600
         Width           =   2415
      End
      Begin VB.OLE OLE1 
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00FFFFC0&
         Class           =   "MPlayer"
         Height          =   255
         Left            =   6360
         OleObjectBlob   =   "Form1.frx":51C1B
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "How Will You Score?"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   6600
         TabIndex        =   4
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Halfpipe Challenge"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   4440
         TabIndex        =   1
         Top             =   720
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdStart_Click()
PName = InputBox("What's Your Name?")
Money = InputBox("What's Your Spending Limit?")

frmFirst.Hide
frmSecond.Show
End Sub
Private Sub cmdQuit_Click()
End
End Sub
