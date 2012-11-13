VERSION 5.00
Begin VB.Form frmhumanisticfollowers 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   8
      Top             =   7920
      Width           =   6015
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   7
      Top             =   7920
      Width           =   6015
   End
   Begin VB.CommandButton Command3 
      Height          =   4095
      Left            =   5640
      Picture         =   "frmhumanisticfollowers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Height          =   4095
      Left            =   10440
      Picture         =   "frmhumanisticfollowers.frx":48D7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Height          =   4095
      Left            =   840
      Picture         =   "frmhumanisticfollowers.frx":ADA7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000040&
      Caption         =   "Viktor Frankl"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   10320
      TabIndex        =   6
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000040&
      Caption         =   "Frederick Perls"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   5280
      TabIndex        =   5
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      Caption         =   "Carl Rogers"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   960
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      Caption         =   "Notable Humanistic Psychologists"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   14535
   End
End
Attribute VB_Name = "frmhumanisticfollowers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmhumanisticfollowers
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: To present major humanistic psychologists. Also, to provide a brief description of their work.
Option Explicit
Private Sub cmdback_Click() 'returns to main humanisic page
frmhumanisticfollowers.Hide
frmhumanistic.Show
End Sub

Private Sub Command1_Click() 'presents a description of Frankl
MsgBox "Viktor Frankl was the founder of logotherapy a supplementary perspective that encourages the patient to find meaning in the world. Many of his ideas were influenced by his time spent in Nazi concentration camps."
End Sub

Private Sub Command2_Click() 'presents a description of Rogers
MsgBox "Carl Rogers founded client centered therapy. He placed an emphasis on the following characteristics during therapy: congruence, unconditional positive regard, and empathy."
End Sub

Private Sub Command3_Click() 'presents a description of Perls
MsgBox "Known as 'Fritz' Perls was the eccentric leader of the Gestalt therapy movement. They focused on dealing with and knowing one's feelings in the present."
End Sub

Private Sub Command4_Click() 'returns to main menu
frmhumanisticfollowers.Hide
frmselectschool.Show
End Sub
