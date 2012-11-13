VERSION 5.00
Begin VB.Form frmEclipse 
   BackColor       =   &H00000080&
   Caption         =   "Eclipse"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEclipseSummary 
      Height          =   3975
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00000080&
      Caption         =   "Click to learn more about Eclipse"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdBookMenu 
      BackColor       =   &H00000080&
      Caption         =   "Click to return to book menu"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to return to main menu"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblEclipse 
      BackColor       =   &H00000080&
      Caption         =   "Eclipse"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   2445
      Left            =   7200
      Picture         =   "frmEclipse.frx":0000
      Top             =   2520
      Width           =   1665
   End
End
Attribute VB_Name = "frmEclipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmEclipse
'Author: Mollie Land
'Date Written: 3/21/09
'Objective: This code is designed to give a summary of the book Eclipse

Private Sub cmdBookMenu_Click()
    'clear the text box
    txtEclipseSummary.Text = ""
    
    'return to book menu, hiding the Eclipse form
    frmBooks.Show
    frmEclipse.Hide
End Sub

Private Sub cmdRead_Click()
    'this button will display the summary of the book in the textbox
    'variables are dimmed publically
    'clear dimmed variables to reset the variable Initial
    Initial = ""
    
    'Open the file to be read
    Open App.Path & "/EclipseSummary.txt" For Input As #1
    
    'Read the file and close it when it has been read
    Do While Not EOF(1)
        Input #1, ReadSummary
        Initial = Initial & ReadSummary & " "
    Loop
    Close (1)
    
    'write the summary that was imported into a
    'text box for the user to see
    txtEclipseSummary.Text = Initial
End Sub

Private Sub cmdReturn_Click()
    'clear the text box
    txtEclipseSummary.Text = ""
    
    'return to main menu, hiding the Eclipse form
    frmStart.Show
    frmEclipse.Hide
End Sub

