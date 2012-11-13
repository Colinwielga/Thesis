VERSION 5.00
Begin VB.Form frmBreakingDawn 
   BackColor       =   &H00000080&
   Caption         =   "Breaking Dawn"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBreakingDawnSummary 
      Height          =   4215
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to learn more about Breaking Dawn"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
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
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
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
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblBreakingDawn 
      BackColor       =   &H00000080&
      Caption         =   "Breaking Dawn"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   2430
      Left            =   7080
      Picture         =   "frmBreakingDawn.frx":0000
      Top             =   2400
      Width           =   1635
   End
End
Attribute VB_Name = "frmBreakingDawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmBreakingDawn
'Author: Mollie Land
'Date Written: 3/21/09
'Objective: This code is designed to give a summary of the book Breaking Dawn

Private Sub cmdBookMenu_Click()
    'clear the text box
    txtBreakingDawnSummary.Text = ""
    
    'return to book menu, hiding Breaking Dawn form
    frmBooks.Show
    frmBreakingDawn.Hide
End Sub

Private Sub cmdRead_Click()
    'this button will display the summary of the book in the textbox
    'variables are dimmed publically
    'clear dimmed variables to reset the variable Initial
    Initial = ""
    
    'Open the file to be read
    Open App.Path & "/BreakingDawnSummary.txt" For Input As #1
    
    'Read the file until all parts have been read and then close it
    Do While Not EOF(1)
        Input #1, ReadSummary
        Initial = Initial & ReadSummary & " "
    Loop
    Close (1)
    
    'write the summary that was imported into a
    'text box for the user to see
    txtBreakingDawnSummary.Text = Initial
End Sub


Private Sub cmdReturn_Click()
    'clear the text box
    txtBreakingDawnSummary.Text = ""
    
    'return to start menu, hiding Breaking Dawn form
    frmStart.Show
    frmBreakingDawn.Hide
End Sub
