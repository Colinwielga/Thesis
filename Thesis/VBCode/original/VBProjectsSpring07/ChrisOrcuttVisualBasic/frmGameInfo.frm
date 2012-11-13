VERSION 5.00
Begin VB.Form frmGameReviews 
   Caption         =   "Game Information"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3135
      Left            =   3720
      ScaleHeight     =   3075
      ScaleWidth      =   7875
      TabIndex        =   4
      Top             =   1080
      Width           =   7935
   End
   Begin VB.CommandButton cmdSearchPlatform 
      Caption         =   "Search Platform"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CommandButton cmdSearchTitle 
      Caption         =   "Search Title"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label lblGameInfo 
      Caption         =   "Looking for a particular game or information? Find it here!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   13695
   End
End
Attribute VB_Name = "frmGameReviews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdReturn_Click()
    frmGameReviews.Hide
    frmSelectWant.Show
End Sub
Private Sub cmdSearchTitle_Click()
    'This program asks a user to input a price range
    'and searches the Retail Store inventory
    'and stops when the first item matching the search is found
    Dim Title(1 To 100) As String
    Dim Platform(1 To 100) As Single
    Dim SearchTitle As Single
    Dim Ctr As Integer
    Dim Found As Boolean
    Dim Pos As Single
    
    SearchTitle = InputBox("Please input a game title to search for", "Search Title")
    
    Open App.Path & "\AllGames.txt" For Input As #1
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Title(Ctr), Platform(Ctr)
    Loop
    Close #1
    
    Pos = 0
    Found = False
    
    Do Until Found = True Or Pos >= Ctr
        Pos = Pos + 1
        If SearchTitle < Platform(Pos) Then
            Found = True
        End If
    Loop
    
        
    If Found = True Then
        picResults.Print Title(Pos); Tab(35); Platform(Pos)
    Else
        picResults.Print "Error"
    End If
End Sub
