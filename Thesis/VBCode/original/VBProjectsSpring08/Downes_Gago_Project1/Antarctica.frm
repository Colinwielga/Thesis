VERSION 5.00
Begin VB.Form Antarctica 
   BackColor       =   &H00FF8080&
   Caption         =   "Antarctica"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   1035
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   15240
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   480
      Picture         =   "Antarctica.frx":0000
      ScaleHeight     =   3345
      ScaleWidth      =   5985
      TabIndex        =   3
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   15975
   End
   Begin VB.CommandButton cmdBacktoMain 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Width           =   2655
   End
   Begin VB.CommandButton cmdAttackofthePenguins 
      Caption         =   "What are those penguins doing?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   5640
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   7320
      Left            =   6960
      Picture         =   "Antarctica.frx":4C63
      Stretch         =   -1  'True
      Top             =   840
      Width           =   8865
   End
End
Attribute VB_Name = "Antarctica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Antarctica.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  This form only gives the user a message box about antarctica
'This form is meant to be a sort of break from all of the endless amounts of data on our other forms

Option Explicit
'Simply gives the user a message box
Private Sub cmdAttackofthePenguins_Click()
    MsgBox ("They Are Trying To Stay Warm. Believe It Or Not, But Antarctica Is WAY Colder Than Minnesota.  It's So Cold Here That Few Animals Spend Their Time Hanging Out On This Continent.  One of The Animals That Do Spend Time On This Continent Are Penguins.")
End Sub
'Hides the Antarctica Fomr and Shows the Main Form
Private Sub cmdBacktoMain_Click()
Antarctica.Hide
Main.Show
End Sub
