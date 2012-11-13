VERSION 5.00
Begin VB.Form frmcognitivefollowers 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
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
      Left            =   7920
      TabIndex        =   4
      Top             =   7440
      Width           =   6015
   End
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
      Left            =   1200
      TabIndex        =   3
      Top             =   7440
      Width           =   6015
   End
   Begin VB.CommandButton Command3 
      Height          =   4095
      Left            =   10680
      Picture         =   "frmcognitivefollowers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Height          =   4095
      Left            =   5760
      Picture         =   "frmcognitivefollowers.frx":4FCD2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Height          =   4095
      Left            =   720
      Picture         =   "frmcognitivefollowers.frx":584F2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Notable Behaviorist/Cognitive Psychologists"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1095
      Left            =   2280
      TabIndex        =   8
      Top             =   600
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "B.F. Skinner"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   7
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Albert Bandura"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   6
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "John Watson"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   4335
   End
End
Attribute VB_Name = "frmcognitivefollowers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmcognitivefollowers
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: To present famous cognitive or behavioral psychologists as well as a brief description of their work.
Option Explicit
Private Sub cmdback_Click() 'returns to cognitive-behavioral page
frmcognitivefollowers.Hide
frmcognitivebehavioral.Show
End Sub

Private Sub Command1_Click() 'provides a description of Watson
MsgBox "Watson was an early founder of this perspective. He is most known for his work with 'Baby Albert' which demonstrated psychologists' ability to manipulate behavior through conditioning."
End Sub

Private Sub Command2_Click() 'provides a description of Bandura
MsgBox "Bandura is a cognitively oriented psychologist. Much of his research has dealt with modeling. His most famous experiments involved exposing children to violent television (a model) and observing their behavior towards bobo dolls."
End Sub

Private Sub Command3_Click() 'provides a description of Skinner
MsgBox "Skinner was a radical behaviorist. He is known for his work with animals demonstrating contingency managment."
End Sub

Private Sub Command4_Click() 'returns to the main menu
frmcognitivefollowers.Hide
frmselectschool.Show
End Sub
