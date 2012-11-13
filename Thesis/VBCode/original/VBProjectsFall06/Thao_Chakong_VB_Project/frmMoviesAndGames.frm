VERSION 5.00
Begin VB.Form frmMovieSale 
   BackColor       =   &H80000001&
   Caption         =   "Movies Available for Sale"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmMoviesAndGames.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewAll 
      Caption         =   "View All Available Movies"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "View List of Movies"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Price"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   3
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtInputTitle 
      Height          =   615
      Left            =   9000
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblInput 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Movie Title:  (note: titles are case sensitive)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   6840
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmMovieSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Planet of Jet Li
'Form Name: frmMovieSale
'Author: Chakong Thao
'Date Written: Wednesday, Nov. 1st
'Form Objective: This form has gone past just listing movies, but
                'offers a text input of any movie title that may
                'have caught the user's interest.  It then searches
                'for a price for that title if it is one of the 18
                'that is for sale.
                
Option Explicit
Private Sub cmdBack_Click() 'This brings user back to General page
    frmMovieSale.Hide
    frmGeneral.Show
End Sub

Private Sub cmdList_Click() 'This helps the user by showing the page of movie listings again to help with the search on this page
    frmMovieSale.Hide
    frmMovies.Show
End Sub

Private Sub cmdMain_Click() 'This brings user back to beginning page
    frmMovieSale.Hide
    frmJetLi.Show
End Sub


Private Sub cmdSearch_Click()   'This reads the array files and will search and determine whether or not the title input by the user is for sale
    Dim Title As String, Price As Single
    Counter = 0
    Open App.Path & "\Prices.txt" For Input As #1
    
    Do Until EOF(1)
        Input #1, Title, Price
        Counter = Counter + 1
        Titles(Counter) = Title
        Prices(Counter) = Price
    Loop
    
    Close #1
    
    Dim X As String, Found As Boolean
    X = txtInputTitle.Text
    Found = False
    Counter = 0
    
    Do Until Found = True Or Counter > 75
        Counter = Counter + 1
        If Titles(Counter) = X Then
            Found = True
        End If
    Loop
    
    If X = "" Then
        MsgBox "Please enter a movie title!", , "Error!"
    ElseIf Found = True Then
        MsgBox "The price for" & " " & X & " " & "is" & " " & FormatCurrency(Prices(Counter)), , "Movie Found!"
    Else
        MsgBox "Sorry, your movie was either not found or not available for sale.", , "Too Bad!"
    End If
    
End Sub

Private Sub cmdViewAll_Click()  'This shows user the page with all movies that are for sale
    frmMovieSale.Hide
    frmAvailability.Show
End Sub
