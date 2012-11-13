VERSION 5.00
Begin VB.Form frmboob 
   BackColor       =   &H00808080&
   Caption         =   "The Boobery"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdboobdrink 
      BackColor       =   &H00FF00FF&
      Caption         =   "What are you drinking?"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00FF00FF&
      Caption         =   "Continue on your tour de st. joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton cmdtalk 
      BackColor       =   &H00FF00FF&
      Caption         =   "Who are you going to talk to?"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Image imageboob 
      Height          =   6795
      Left            =   840
      Picture         =   "frmboob.frx":0000
      Top             =   840
      Width           =   9060
   End
   Begin VB.Label lblboob 
      BackColor       =   &H00808080&
      Caption         =   "               Welcome to the Boobery"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmboob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdboobdrink_Click()

    frmboobdrink.Show
    frmboob.Hide

End Sub

    'Project name:  Tour De St. Joe
    'Form:  frmboob, "The Boobery"
    'Author:  Brooke
    'Date:  3/12/08
    'Objective:  One of the places to stop.  Using V.B. tools to create a fun/realistic setting.
    


Private Sub cmdleave_Click()

    frmjoetown.Show
    frmboob.Hide

End Sub

Private Sub cmdtalk_Click()

    frmtalkto.Show
    frmboob.Hide

End Sub


