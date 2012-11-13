VERSION 5.00
Begin VB.Form frmTwilight 
   BackColor       =   &H00000080&
   Caption         =   "Twilight"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTwilightSummary 
      Height          =   3855
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00000080&
      Caption         =   "Click to learn more about Twilight"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdBookMenu 
      BackColor       =   &H00000080&
      Caption         =   "Click to return to the book menu"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to return to the main menu"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblTwilight 
      BackColor       =   &H00000080&
      Caption         =   "Twilight"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   7440
      Picture         =   "frmTwilight.frx":0000
      Top             =   2280
      Width           =   1620
   End
End
Attribute VB_Name = "frmTwilight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmTwilight
'Author: Mollie Land
'Date Written: 3/21/09
'Objective: This code is designed to give a summary of the book Twilight

Private Sub cmdBookMenu_Click()
    'clear the text box
    txtTwilightSummary.Text = ""
    
    'return to book menu, hiding Twilight form
    frmBooks.Show
    frmTwilight.Hide
End Sub

Private Sub cmdRead_Click()
    'this button will display the summary of the book in the textbox
    'variables are dimmed publically
    'clear dimmed variables to reset the variable Initial
    Initial = ""
    
    'Open the file to be read
    Open App.Path & "/TwilightSummary.txt" For Input As #1
    
    'Read the file that was imported until are parts have been read, then close the file
    Do While Not EOF(1)
        Input #1, ReadSummary
        Initial = Initial & ReadSummary & " "
    Loop
    Close (1)
    
    
    'write the summary that was imported into a
    'text box for the user to see
    txtTwilightSummary.Text = Initial
End Sub

Private Sub cmdReturn_Click()
    'clear the text box
    txtTwilightSummary.Text = ""
    
    'return to start menu, hiding Twilight form
    frmStart.Show
    frmTwilight.Hide
End Sub
