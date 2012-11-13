VERSION 5.00
Begin VB.Form frmDescription2 
   BackColor       =   &H00000000&
   Caption         =   "Types of Pedophiles"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000FF00&
      Caption         =   "Back to case files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9720
      Width           =   2295
   End
   Begin VB.PictureBox picpre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   7800
      ScaleHeight     =   7455
      ScaleWidth      =   7455
      TabIndex        =   3
      Top             =   1920
      Width           =   7455
   End
   Begin VB.CommandButton cmdpre 
      BackColor       =   &H000080FF&
      Caption         =   "Click to review prefertial molesters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.PictureBox picsit 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   120
      ScaleHeight     =   7455
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   1920
      Width           =   7575
   End
   Begin VB.CommandButton cmdsit 
      BackColor       =   &H000080FF&
      Caption         =   "Click to review situational molesters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2160
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmDescription2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBack_Click()
'Takes the user back to previous form
    frmDescription2.Hide
    frmCasefiles.Show
    
End Sub

Private Sub cmdpre_Click()
'Declare all my variables for this button
Dim pos As Integer, ctr As Integer
    
    picpre.Cls 'Clears out the picture box of any unwanted material
 'Open a file that i created and store it as #5
    Open App.Path & "\sitpre.txt" For Input As #5
    Do Until EOF(5) 'I want to read the file until it has read the entire thing
        pos = pos + 1 'keeps track of how many lines i am reading
        Input #5, prepedo(pos) 'stores them in this variable
    Loop 'I loop until the file has been completely read
    Close #5 'Close the file as to not interrupt the readings from other files
    
    For ctr = 21 To 37 'the range that i want to have the program print
        picpre.Print prepedo(ctr) 'the result of the printing amount
    Next ctr 'continues the process until the ctr has reached its total amount
        
End Sub

Private Sub cmdsit_Click()
'Declare all my variables for this button
Dim pos As Integer, ctr As Integer
    
    picsit.Cls 'clears out the picture box of any unwanted material
    'Open a file that i created and store it as #4
        Open App.Path & "\sitpre.txt" For Input As #4
        Do Until EOF(4) 'Read the file until is has read the entire file
            pos = pos + 1 'keeps track of how many lines i am reading
            Input #4, sitpedo(pos) 'stores them in this variable
        Loop 'Loops until the file has been completely read
        Close #4 'Close the file so it doesn't mess up the reading of another file
        
    For ctr = 1 To 19 'the range that i want to have the program print
        picsit.Print sitpedo(ctr) 'the result of the printing amount
    Next ctr 'continues the process until the ctr has reached its total amount
End Sub
