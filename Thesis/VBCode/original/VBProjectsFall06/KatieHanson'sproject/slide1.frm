VERSION 5.00
Begin VB.Form Title 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCite 
      BackColor       =   &H000000FF&
      Caption         =   "Citation"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H000000FF&
      Caption         =   "Purchase Movie"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdClick 
      Height          =   1935
      Left            =   6360
      Picture         =   "slide1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdBreakup 
      Height          =   1935
      Left            =   4440
      Picture         =   "slide1.frx":1491
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdPirates 
      Height          =   1935
      Left            =   2640
      Picture         =   "slide1.frx":27DF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCode 
      Height          =   1935
      Left            =   8280
      Picture         =   "slide1.frx":403F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdTruth 
      Height          =   1935
      Left            =   8280
      Picture         =   "slide1.frx":5674
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdCars 
      Height          =   1935
      Left            =   4440
      Picture         =   "slide1.frx":681C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdGeorge 
      BackColor       =   &H80000007&
      Height          =   1935
      Left            =   6360
      Picture         =   "slide1.frx":80F7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdlake 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2640
      MaskColor       =   &H00C0FFFF&
      Picture         =   "slide1.frx":9888
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H000000FF&
      Caption         =   $"slide1.frx":AB13
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2520
      TabIndex        =   10
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Label lblMovies 
      BackColor       =   &H00000000&
      Caption         =   "New Movies"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   24
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: Title
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to click on the movies pictured and find out more information about the movies and also purchase the movies listed.

Option Explicit
'Move to the movie "the breakup" form
Private Sub cmdBreakup_Click()
    breakup.Show
    Title.Hide
End Sub
'Move to the movie "Cars" form
Private Sub cmdCars_Click()
    Cars.Show
    Title.Hide
End Sub

'Move to the form that shows where i got my information
Private Sub cmdCite_Click()
    Cite.Show
    Title.Hide
End Sub

'Move to the movie "Click" form
Private Sub cmdClick_Click()
    Click.Show
    Title.Hide
End Sub
'Move to the movie "The Davinci Code" form
Private Sub cmdCode_Click()
    Code.Show
    Title.Hide
End Sub
'Move to the movie "Curious George" form
Private Sub cmdGeorge_Click()
    George.Show
    Title.Hide
End Sub
'Move to the movie "The Lakehouse" form
Private Sub cmdlake_Click()
    lakehouse.Show
    Title.Hide
End Sub
'Move to the movie "The Pirates of the Caribbean" form
Private Sub cmdPirates_Click()
    Pirates.Show
    Title.Hide
End Sub
'Move to the Purchase form to purchase the movies
Private Sub cmdPurchase_Click()
    Purchase.Show
    Title.Hide
End Sub
'move to the movie "An inconvient truth" form
Private Sub cmdTruth_Click()
    Truth.Show
    Title.Hide
End Sub










