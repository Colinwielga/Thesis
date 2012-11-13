VERSION 5.00
Begin VB.Form frmNewMoon 
   BackColor       =   &H00000080&
   Caption         =   "New Moon"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNewMoonSummary 
      Height          =   3855
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to learn more about New Moon"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdBookMenu 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
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
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblNewMoon 
      BackColor       =   &H00000080&
      Caption         =   "New Moon"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2430
      Left            =   7200
      Picture         =   "frmNewMoon.frx":0000
      Top             =   2520
      Width           =   1710
   End
End
Attribute VB_Name = "frmNewMoon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmNewMoon
'Author: Mollie Land
'Date Written: 3/21/09
'Objective: This code is designed to give a summary of the book New Moon

Private Sub cmdBookMenu_Click()
    'clear the text box
    txtNewMoonSummary.Text = ""
    
    'return to book menu, hiding the New Moon form
    frmBooks.Show
    frmNewMoon.Hide
End Sub


Private Sub cmdRead_Click()
    'this button will display the summary of the book in the textbox
    'variables are dimmed publically
    'clear dimmed variables to reset the variable Initial
    Initial = ""
    
    'Open the file to be read
    Open App.Path & "/NewMoonSummary.txt" For Input As #1
     
    'read the file and close it when completed
    Do While Not EOF(1)
        Input #1, ReadSummary
        Initial = Initial & ReadSummary & " "
    Loop
    Close (1)
    
    'write the summary that was imported into a
    'text box for the user to see
    txtNewMoonSummary.Text = Initial
End Sub


Private Sub cmdReturn_Click()
    'clear the text box
    txtNewMoonSummary.Text = ""
    
    'return to start menu, hiding the New Moon form
    frmStart.Show
    frmNewMoon.Hide
End Sub

