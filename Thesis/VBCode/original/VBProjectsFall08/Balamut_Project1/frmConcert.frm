VERSION 5.00
Begin VB.Form frmConcert 
   BackColor       =   &H000000FF&
   Caption         =   "Weezer in Concert"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPic9 
      Caption         =   "Picture #9"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   15
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdQuestion 
      Caption         =   "Question for YOU!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdTour 
      Caption         =   "Tour page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdPic6 
      Caption         =   "Picture #6"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdPic7 
      Caption         =   "Picture #7"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdPic8 
      Caption         =   "Picture #8"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdPic10 
      Caption         =   "Picture #10"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit This Rad Program"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPic2 
      Caption         =   "Picture #2"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdPic3 
      Caption         =   "Picture #3"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdPic4 
      Caption         =   "Picture #4"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdPic5 
      Caption         =   "Picture #5"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdPic1 
      Caption         =   "Picture #1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      Height          =   6735
      Left            =   240
      ScaleHeight     =   6675
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   1680
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Weezer's Troublemaker Tour Concert at the XCel Energy Center on 10/3/08"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmConcert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Weezer
'Form Name: frmConcert.frm
'Author: Emily Balamut
'Date Written: 10/30/08
'Objective: This form allows the user to click on buttons and see pictures from
'the Weezer concert that I myself attended. It also has a question for the user
'to answer.
Option Explicit

Private Sub cmdMain_Click()
    frmConcert.Hide
    frmBeginning.Show
End Sub

Private Sub cmdPic1_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert2.jpg")
End Sub

Private Sub cmdPic10_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert3.jpg")
End Sub

Private Sub cmdPic2_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert4.jpg")
End Sub

Private Sub cmdPic3_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert11.jpg")
End Sub

Private Sub cmdPic4_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert13.jpg")
End Sub

Private Sub cmdPic5_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert12.jpg")
End Sub

Private Sub cmdPic6_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert7.jpg")
End Sub

Private Sub cmdPic7_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert10.jpg")
End Sub

Private Sub cmdPic8_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert8.jpg")
End Sub

Private Sub cmdPic9_Click()
    picResults.Picture = LoadPicture(App.Path & "\concert9.jpg")
End Sub

Private Sub cmdQuestion_Click()
    Dim Question As String
    Dim Number As Integer
    
    Question = InputBox("Yes or No: Have you ever seen Weezer in concert?", , "Question")
    
    If LCase(Question) = LCase("Yes") Then
        Number = InputBox("How many times?", , "Times")
        picResults.Cls
        Select Case Number
            Case Is = 1
                MsgBox "That's how many times I've seen them, too!", , "Me too!"
            Case Is = 2
                MsgBox "Twice as many as me! That's awesome!", , "Cool"
            Case Is = 3
                MsgBox "Wow! You must have been a fan for a very long time!", , "Neato!"
            Case Is = 4
                MsgBox "You are really lucky! I wish I have seen Weezer 4 times!", , "Holy Cow!"
            Case Is >= 5
                MsgBox "You are the only person I know that has seen Weezer live so many times. I am in awe!", , "Awesome!"
        End Select
    End If
    
    If LCase(Question) = LCase("No") Then
        MsgBox "Oh, that's really too bad. You should really check out a concert at their next tour!", , "Sad Day"
    End If
            
End Sub

Private Sub cmdQuit_Click()
MsgBox "Thanks for rocking out with Weezer, " & UserName & "! See you later!", , "Bye!"
End
End Sub

Private Sub cmdTour_Click()
    frmConcert.Hide
    frmSchedule.Show
End Sub
