VERSION 5.00
Begin VB.Form frmarray 
   Caption         =   "Form1"
   ClientHeight    =   11205
   ClientLeft      =   3435
   ClientTop       =   2415
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   Picture         =   "frmarray.frx":0000
   ScaleHeight     =   11205
   ScaleWidth      =   14685
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0080FF80&
      Caption         =   "Back to Neverland!"
      Height          =   975
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9600
      Width           =   2535
   End
   Begin VB.CommandButton cmdyear 
      BackColor       =   &H0080FFFF&
      Caption         =   "View by Year"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton cmdalpha 
      BackColor       =   &H0080FFFF&
      Caption         =   "View Alphabetically"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdread 
      BackColor       =   &H0080FFFF&
      Caption         =   "Read the File"
      Height          =   855
      Left            =   6840
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.PictureBox picresults 
      Height          =   9855
      Left            =   1080
      ScaleHeight     =   9795
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "What are Disney's Masterpiece Movies?"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   6000
      TabIndex        =   6
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label lblmovie 
      Caption         =   $"frmarray.frx":3B3EE
      Height          =   615
      Left            =   6840
      TabIndex        =   1
      Top             =   8760
      Width           =   4215
   End
End
Attribute VB_Name = "frmarray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Wonderful World of Disney
'form Home
'Kate Hendrickx
'February 2010
' Objective: this form lists all the movies that Disney has classified as masterpieces, and are animated.
' the user has the option to view the list by chronological year or by alphabetical title.
Option Explicit
Dim alphabet(1 To 45) As String, year(1 To 45) As Single, CTR As Single
Dim ctr2 As Single, W As Integer, D As Integer


Private Sub cmdread_Click()
'opening the file
Open App.Path & "\movielist.txt" For Input As #1
picresults.Print
picresults.Print "Movie Title"; Tab(40); "Year Released"
picresults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
picresults.Print

'reading the array
CTR = 0
Do While Not EOF(1)
CTR = CTR + 1
Input #1, alphabet(CTR), year(CTR)
picresults.Print alphabet(CTR); Tab(45); year(CTR)
Loop
Close #1

'Enabling other buttons
cmdalpha.Enabled = True
cmdyear.Enabled = True
End Sub
Private Sub cmdalpha_Click()
'declaring the variables
Dim Pass As Integer, Pos As Integer, Temp As String

'clearing the picture box of previous info
picresults.Cls

'setting up the header
picresults.Print
picresults.Print "Movie Title"; Tab(40); "Year Released"
picresults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
picresults.Print

'organizing the info alphabetically by bubble sort
For Pass = 1 To CTR - 1
For Pos = 1 To CTR - Pass
    If alphabet(Pos) > alphabet(Pos + 1) Then
     Temp = alphabet(Pos)
     alphabet(Pos) = alphabet(Pos + 1)
     alphabet(Pos + 1) = Temp
    End If
    Next Pos
    Next Pass
    
' printing the sorted list
For W = 1 To CTR
picresults.Print alphabet(W); Tab(45); year(W)
Next W

End Sub

Private Sub cmdBack_Click()
frmarray.Hide
frmhome.Show
End Sub

Private Sub cmdyear_Click()
'declaring the variables
Dim Temp2 As Integer, Pass2 As Integer, Pos2 As Integer

'clearing the picture box from previous info
picresults.Cls

'setting up the header
picresults.Print
picresults.Print "Year Released"; Tab(20); "Movie Title"
picresults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
picresults.Print

'organizing the info chronologically by bubble sort
For Pass2 = 1 To CTR - 1
For Pos2 = 1 To CTR - Pass2
    If year(Pos2) > year(Pos2 + 1) Then
     Temp2 = year(Pos2)
     year(Pos2) = year(Pos2 + 1)
     year(Pos2 + 1) = Temp2
    End If
    Next Pos2
    Next Pass2
    
' printing the sorted list
For D = 1 To CTR
picresults.Print year(D); Tab(20); alphabet(D)
Next D
End Sub

