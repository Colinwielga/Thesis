VERSION 5.00
Begin VB.Form frmpsychodynamic 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   Picture         =   "frmpsychodynamic.frx":0000
   ScaleHeight     =   4.07357e7
   ScaleMode       =   0  'User
   ScaleWidth      =   4.2874e8
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9120
      TabIndex        =   3
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdquizes 
      Caption         =   "Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      TabIndex        =   2
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdmajortheories 
      Caption         =   "Major Theories and Important Techniques"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      TabIndex        =   1
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmddisciples 
      Caption         =   "Disciples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Psychodynamic"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmpsychodynamic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmpsychodynamic
'Author: Calvin Pipenhagen
'Date Written: March 9, 2008
'Objective: This form holds the links to all information and activities related to the Psychodynamic orientation.
           'The user clicks on a button and is transferred to a new form.
Option Explicit

Private Sub cmddisciples_Click() 'displays a form showing famous psychodynamic psychologists
frmdisciples.Show
frmpsychodynamic.Hide
End Sub

Private Sub cmdmajortheories_Click() 'displays a form of important theories
frmpsychodynamic.Hide
frmmajortheories.Show
End Sub

Private Sub cmdquizes_Click() ' displays a quiz
frmpsychodynamic.Hide
frmpsychodynamicquiz.Show
End Sub

Private Sub Command4_Click() 'returns to main menu
frmselectschool.Show
frmpsychodynamic.Hide
End Sub


