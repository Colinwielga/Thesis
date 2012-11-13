VERSION 5.00
Begin VB.Form frmNikeBaseball 
   BackColor       =   &H80000006&
   Caption         =   "NikeBaseball"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   Picture         =   "frmNikeBaseball.frx":0000
   ScaleHeight     =   8430
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   6015
      Left            =   7920
      ScaleHeight     =   5955
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go Back to the Store Home"
      Height          =   1215
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go Back to Nike Store"
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortPrice 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort By the Price"
      Height          =   1335
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortName 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort By Name of Item"
      Height          =   1335
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0000FFFF&
      Caption         =   "Read Data"
      Height          =   1335
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   1215
      Left            =   8400
      OleObjectBlob   =   "frmNikeBaseball.frx":68322
      SourceDoc       =   "M:\CS130\AndrewEisingerVBproject\ESPN_-_Baseball_Tonight.mp3"
      TabIndex        =   7
      Top             =   6720
      Width           =   2655
   End
End
Attribute VB_Name = "frmNikeBaseball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' NikeBaseball
' Andrew Eisinger
' 3/19/09
'This program reads a file
'This program then based on the file can display and then sort by price and alphabetically
Dim BaseballItems(1 To 50) As String, BaseballCosts(1 To 50) As Single


Private Sub cmdBackHome_Click()
frmStoreHome.Show
frmNikeBaseball.Hide
End Sub

Private Sub cmdGoBack_Click()
frmNike1.Show
frmNikeBaseball.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()
picResults.Cls
Open App.Path & "\NikeBaseball.txt" For Input As #1
CTR = 0
picResults.Print "Baseball Items"; Tab(40); "Baseball Costs"
picResults.Print "*****************************************************************************"
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, BaseballItems(CTR), BaseballCosts(CTR)
    picResults.Print BaseballItems(CTR); Tab(40); FormatCurrency(BaseballCosts(CTR))
Loop
Close #1
End Sub

Private Sub cmdSortName_Click()
'This button sorts the arrays by the first names of the products
Dim J As Single, TempBaseballItems As String, TempBaseballCosts As Single
    'Clears the previous results
    picResults.Cls
    
    'Prints headings so it is easier for the user to understand
    'the results.
    picResults.Print "Baseball Item"; Tab(40); "Baseball Cost"
    picResults.Print "**************************************************************************"

    
    'Code to sort the two parralel arrays by the first name of the product
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If BaseballItems(Pos) > BaseballItems(Pos + 1) Then
                TempBaseballItems = BaseballItems(Pos)
                BaseballItems(Pos) = BaseballItems(Pos + 1)
                BaseballItems(Pos + 1) = TempBaseballItems
                TempBaseballCosts = BaseballCosts(Pos)
                BaseballCosts(Pos) = BaseballCosts(Pos + 1)
                BaseballCosts(Pos + 1) = TempBaseballCosts
            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print BaseballItems(J); Tab(40); FormatCurrency(BaseballCosts(J))
    Next
End Sub

Private Sub cmdSortPrice_Click()
'This button sorts the arrays by the price of the products
Dim J As Single, TempBaseballItems As String, TempBaseballCosts As Single
    'Clears the previous results
    picResults.Cls
    
    'Prints headings so it is easier for the user to understand
    'the results.
    picResults.Print "Baseball Item"; Tab(40); "Baseball Cost"
    picResults.Print "**************************************************************************"

    
    'Code to sort the two parralel arrays by the price of the product
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If BaseballCosts(Pos) > BaseballCosts(Pos + 1) Then
                TempBaseballCosts = BaseballCosts(Pos)
                BaseballCosts(Pos) = BaseballCosts(Pos + 1)
                BaseballCosts(Pos + 1) = TempBaseballCosts
                TempBaseballItems = BaseballItems(Pos)
                BaseballItems(Pos) = BaseballItems(Pos + 1)
                BaseballItems(Pos + 1) = TempBaseballItems
            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print BaseballItems(J); Tab(40); FormatCurrency(BaseballCosts(J))
    Next
End Sub
