VERSION 5.00
Begin VB.Form frmJetLi 
   BackColor       =   &H80000000&
   Caption         =   "Planet of Jet Li"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2280
      TabIndex        =   1
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "  WELCOME TO PLANET OF JET LI"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   7215
   End
End
Attribute VB_Name = "frmJetLi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Planet of Jet Li
'Form Name: frmJetLi
'Author: Chakong Thao
'Date Written: Sunday, Oct. 29th
'Form Objective: This form gives the program a title page and a limited
                'number of command buttons to start off easy.
'Overall Objective:  The objective is to inform users who have never
                    'heard of the great actor, Jet Li.  It may also catch
                    'some interest in users that may not know where
                    'to start, and that is the reason for the movie
                    'listings.  It gives the users a list of many of
                    'his movies that they may not have seen before.
                    'This gives them the chance to actually go out
                    'and buy a Jet Li film of their interest.

Option Explicit

Private Sub cmdExit_Click() 'This exits the entire program
    End
End Sub

Private Sub cmdStart_Click()    'This button hides this current form and opens up the one named General, where the user will find more commands
    frmJetLi.Hide
    frmGeneral.Show
End Sub
