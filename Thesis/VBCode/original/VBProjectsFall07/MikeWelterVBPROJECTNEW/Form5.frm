VERSION 5.00
Begin VB.Form frmFifth 
   Caption         =   "Are You Ready"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   10335
      Left            =   -720
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   10275
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   -720
      Width           =   11895
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Start Your First Run"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3600
         Width           =   9375
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FFFFC0&
         Caption         =   "I Changed My Mind... I Want Out"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4560
         Width           =   9375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "But Do You Have What It Takes To Win?"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   960
         TabIndex        =   2
         Top             =   2160
         Width           =   9975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "So You've Got Your Gear....."
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   2880
         TabIndex        =   1
         Top             =   1080
         Width           =   10095
      End
   End
End
Attribute VB_Name = "frmFifth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNext_Click()

frmFifth.Hide
frmSixth.Show

End Sub

Private Sub cmdQuit_Click()

End

End Sub
