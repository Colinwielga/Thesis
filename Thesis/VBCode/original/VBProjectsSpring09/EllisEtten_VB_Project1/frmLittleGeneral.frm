VERSION 5.00
Begin VB.Form frmLittleGeneral 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   4575
   End
   Begin VB.CommandButton cmdGotoTotals 
      BackColor       =   &H00008000&
      Caption         =   "Go to Totals"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   4575
   End
   Begin VB.CommandButton cmdEventsAttended 
      BackColor       =   &H00008000&
      Caption         =   "Find Attened Events"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   4575
   End
   Begin VB.CommandButton cmdListGeneral 
      BackColor       =   &H00008000&
      Caption         =   "list events"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   5175
      Left            =   6000
      ScaleHeight     =   5115
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lblBigGeneral 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "All those other fun points"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   9255
   End
End
Attribute VB_Name = "frmLittleGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
