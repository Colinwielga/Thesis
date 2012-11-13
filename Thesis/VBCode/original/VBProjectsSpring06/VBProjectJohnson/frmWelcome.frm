VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00FF0000&
   Caption         =   "2006 St. Johns Housing Draft"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H000000FF&
      Caption         =   "Enter The Draft Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00FF0000&
      Caption         =   "Project Created By:                 Kyle Johnson"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4080
      TabIndex        =   3
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00FF0000&
      Caption         =   "         Welcome To           St. John's University           Housing Draft                         2006"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. Johns Housing Project
' Welcome Form
' Written By Kyle Johnson
' 3/22/06
' This form serves as a welcome page and notifies the user as to what
' this program will accomplish
' It allows the User to either enter the draft, or
' The overall purpose of this project is to replicate the
' st. johns housing lottery process.
' It allows users to interact within the program to learn more about
' the housing options, and ultimately choose a house
' The program saves all the housing choices, then allows the user to
' output that information to a text file to be used by faculty.


Private Sub cmdEnter_Click()
    'navigates from the welcome page to the draft page
    frmDraft.Visible = True
    frmWelcome.Visible = False

    'informs user of first step in the draft process
    MsgBox "Start By Loading The File", , "Step 1"


End Sub

Private Sub cmdExit_Click()
    'exits the program
    End
End Sub
