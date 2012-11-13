VERSION 5.00
Begin VB.Form FormRUNS 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtitleruns 
      BackColor       =   &H008080FF&
      Caption         =   "To Title"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   960
      ScaleHeight     =   5595
      ScaleWidth      =   7755
      TabIndex        =   1
      Top             =   1200
      Width           =   7815
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H0000FFFF&
      Caption         =   "PRESS HERE TO SEE RECOMMENDED RUNS FOR YOU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   9495
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   1680
      Picture         =   "FormRuns.frx":0000
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   6540
   End
   Begin VB.Shape Shape9 
      DrawMode        =   16  'Merge Pen
      Height          =   2655
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   -1320
      Width           =   2655
   End
   Begin VB.Shape Shape8 
      DrawMode        =   16  'Merge Pen
      Height          =   3735
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Shape Shape7 
      DrawMode        =   16  'Merge Pen
      Height          =   735
      Left            =   600
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      DrawMode        =   16  'Merge Pen
      Height          =   1215
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   0
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      DrawMode        =   16  'Merge Pen
      Height          =   2655
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      DrawMode        =   16  'Merge Pen
      Height          =   495
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   495
   End
   Begin VB.Shape Shape3 
      DrawMode        =   16  'Merge Pen
      Height          =   1095
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      DrawMode        =   16  'Merge Pen
      Height          =   2655
      Left            =   -1440
      Shape           =   3  'Circle
      Top             =   0
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      DrawMode        =   16  'Merge Pen
      Height          =   1335
      Left            =   -360
      Shape           =   3  'Circle
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "FormRUNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SKI TRIP
'SKI RUNS
'HOLLIS FRITTS
'8-18
'THIS IS INFORMATION ON WHAT SKI RUNS TO ATTEMPT


Option Explicit

Private Sub cmdCalculate_Click()
'dim variables
Dim Difficulty As Single
'make an inputBox to determine the users difficulty level
Difficulty = InputBox("Enter your difficulty level (1-3)                                                                                                                                       1 - Beginner      2 - Intermediate      3 - Expert")
    'make If statements to determine what runs the user should try
    'after input is recieved print image and text to a picture box
        If Difficulty = 1 Then
           picResults.Cls
           picResults.Picture = LoadPicture(App.Path & "\EASYRUN.bmp")
           picResults.Print "You Should Try Runs Like This"
          
        End If
        
        
        If Difficulty = 2 Then
            picResults.Cls
            picResults.Picture = LoadPicture(App.Path & "\INTERRUN.bmp")
            picResults.Print "You Should Try Runs Like This"
        End If
        
        
        If Difficulty = 3 Then
            picResults.Cls
            picResults.Picture = LoadPicture(App.Path & "\EXPERTRUN.bmp")
            picResults.Print "You Should Try Runs Like This"
        End If
        
        
        If Difficulty Then
        Else
            picResults.Print "PLEASE ENTER A VALID NUMBER"
        End If


End Sub

Private Sub cmdtitleruns_Click()
FormRUNS.Hide
Title.Show
End Sub

