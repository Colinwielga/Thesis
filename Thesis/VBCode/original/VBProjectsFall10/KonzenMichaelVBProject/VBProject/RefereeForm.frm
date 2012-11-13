VERSION 5.00
Begin VB.Form frmreferee 
   Caption         =   "Referees"
   ClientHeight    =   10890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   Picture         =   "Referee Form.frx":0000
   ScaleHeight     =   10890
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   5
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton cmdlastname 
      Caption         =   "Put Referees in Alphabetical order by last name"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton cmdalphabetize 
      Caption         =   "Put Referees in Alphabetical Order by First Name"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   3
      Top             =   5760
      Width           =   3975
   End
   Begin VB.CommandButton cmdshowref 
      Caption         =   "View Referee for the UFC"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      Picture         =   "Referee Form.frx":21C7F
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   5280
      ScaleHeight     =   3675
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   6960
      Width           =   4455
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Go Back to Main Screen"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      TabIndex        =   0
      Top             =   9600
      Width           =   2295
   End
End
Attribute VB_Name = "frmreferee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdalphabetize_Click()
Dim pass As Integer, pos As Integer, temp As String, ctr As Integer, firstname(1 To 100) As String
Dim I As Integer, lastname(1 To 100) As String, temp2 As String

    
Open App.Path & "\Referees.txt" For Input As #1 'opening file
    ctr = 0
Do While Not EOF(1)
ctr = ctr + 1
Input #1, firstname(ctr), lastname(ctr)
Loop

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass 'putting data into a different order
     If firstname(pos) > firstname(pos + 1) Then
        temp = firstname(pos)
        firstname(pos) = firstname(pos + 1)
        firstname(pos + 1) = temp
        
        temp2 = lastname(pos)
        lastname(pos) = lastname(pos + 1)
        lastname(pos + 1) = temp2
     End If
    Next pos
    Next pass
For I = 1 To ctr
    picresults.Print firstname(I), lastname(I) 'printing in the order I declared
Next I

End Sub

Private Sub cmdclear_Click()
picresults.Cls
End Sub 'clear the picture box

Private Sub cmdgoback_Click()
frmreferee.Hide 'hide this frame and show the main
frmmainscreen.Show
End Sub

Private Sub cmdlastname_Click()
Dim pass As Integer, pos As Integer, temp As String, ctr As Integer, firstname(1 To 100) As String
Dim I As Integer, lastname(1 To 100) As String, temp2 As String

    
Open App.Path & "\Referees.txt" For Input As #1 'opened the text file
    ctr = 0
Do While Not EOF(1)
ctr = ctr + 1
Input #1, firstname(ctr), lastname(ctr) 'defined what is in the text file
Loop

For pass = 1 To ctr - 1 'arranging whats in the file by lastname
    For pos = 1 To ctr - pass
     If lastname(pos) > lastname(pos + 1) Then
        temp = firstname(pos)
        lastname(pos) = lastname(pos + 1)
        lastname(pos + 1) = temp
        
        temp2 = lastname(pos)
        lastname(pos) = lastname(pos + 1)
        lastname(pos + 1) = temp2
     End If
    Next pos 'moves on to next thing in file
    Next pass
For I = 1 To ctr
    picresults.Print lastname(I), firstname(I) 'printing what is in the file, but in the order i wanted
Next I
End Sub

Private Sub cmdshowref_Click()
Dim referee As String, ctr As Integer, firstname As String, lastname As String
    ctr = 0
Open App.Path & "\Referees.txt" For Input As #1 'opeining a text file

Do While Not EOF(1)
Input #1, firstname, lastname 'defining whats in the file
    ctr = ctr + 1
    picresults.Print firstname, lastname 'prints out what is in the file
Loop
Close #1
End Sub
