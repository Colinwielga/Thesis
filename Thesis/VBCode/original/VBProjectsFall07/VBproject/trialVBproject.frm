VERSION 5.00
Begin VB.Form frmDoll 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   Picture         =   "trialVBproject.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   13035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDonate 
      Caption         =   "DONATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   10110
      Picture         =   "trialVBproject.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdShowdoll 
      Caption         =   "Meet Thelma!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   750
      Picture         =   "trialVBproject.frx":1103
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Let's See!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4470
      Picture         =   "trialVBproject.frx":1F03
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   4095
   End
   Begin VB.CommandButton cmdbottom 
      Caption         =   "Bottoms!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7590
      Picture         =   "trialVBproject.frx":266E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "BYE!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9990
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "YES!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7830
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "Name!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4830
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2670
      ScaleHeight     =   675
      ScaleWidth      =   7635
      TabIndex        =   1
      Top             =   240
      Width           =   7695
   End
   Begin VB.CommandButton cmdshirt 
      Caption         =   "Shirt!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4830
      Picture         =   "trialVBproject.frx":33FD
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   2175
   End
End
Attribute VB_Name = "frmDoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdbottom_Click()
'Makes different forms visible
frmbottoms.Show
frmDoll.Hide
End Sub

Private Sub cmdDonate_Click()
'Makes different forms visible
frmDoll.Hide
frmDonate.Show
End Sub

Private Sub cmdDone_Click()
'Makes different forms visible
frmDoll.Hide
frmResult.Show

    


End Sub

Private Sub cmdName_Click()
Dim dollname As String
'Asks for Input from the User and prints a question with that input in it

dollname = InputBox("What is your name?")
picResults.Print dollname; ",  Will you help me get dressed for a day with my friends?"

End Sub

Private Sub cmdQuit_Click()
'Exits the program
End
End Sub

Private Sub cmdshirt_Click()
'Opens a new form
frmDoll.Hide
frmshirts.Show
picResults.Print "Now will you help me pick pants?"
End Sub

Private Sub cmdShowdoll_Click()
'Opens a new form
frmDoll.Hide
frmThelma.Show

End Sub

Private Sub cmdYes_Click()
'Clears the picturebox and prints a new comment
picResults.Cls
picResults.Print "Great! Click on the Shirt button and then choose a color!"
End Sub


