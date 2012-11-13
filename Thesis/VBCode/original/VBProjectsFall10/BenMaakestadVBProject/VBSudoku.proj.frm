VERSION 5.00
Begin VB.Form frmName 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11070
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   600
      ScaleHeight     =   915
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   8280
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit..."
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Name."
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "Enter Name!"
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   2910
      Left            =   6840
      Picture         =   "VB Sudoku.proj.frx":0000
      Top             =   720
      Width           =   3885
   End
   Begin VB.Image Image1 
      Height          =   3765
      Left            =   120
      Picture         =   "VB Sudoku.proj.frx":1343
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   " Welcome to VB Sudoku!"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form allows the user to enter their name and change their name.
'Almost anything may be declared as the users name

Private Sub cmdChange_Click()
    Dim hold As String
    
    
    hold = InputBox("Enter a new name.", "New Name") 'takes the users name length to make sure somethig was entered
    If Len(hold) > 0 Then
        User = hold
    End If
    picResults.Cls
    picResults.Print User
End Sub

Private Sub cmdName_Click() 'user enters their name, instructions
    User = InputBox("Enter your name... when finished, submit to procede!", "Enter Name")
    MsgBox "Welcome to VB Sudoku " & User & "! If this is the incorrect name, click on the option below.", , "Welcome"
    picResults.Print User
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSubmit_Click()
    If Len(User) > 0 Then   'enforces the name length, as long as something was entered the user may procede
        frmName.Hide
        frmOption.Show
    Else
        MsgBox "The program would like to know who you are before you procede."
    End If
        
End Sub
