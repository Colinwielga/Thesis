VERSION 5.00
Begin VB.Form frmNikeFootball 
   Caption         =   "NikeFootball"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   Picture         =   "frmNikeFootball.frx":0000
   ScaleHeight     =   9030
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdGoBackHome 
      BackColor       =   &H00FF0000&
      Caption         =   "Go Back to Store Home"
      Height          =   1095
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Go Back to Nike Store"
      Height          =   1095
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortbyPrice 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort by Price"
      Height          =   1095
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortbyName 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort by Item Name"
      Height          =   1095
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdReadData 
      BackColor       =   &H00FF0000&
      Caption         =   "Read Data"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000C&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4875
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   4080
      Width           =   5055
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   1455
      Left            =   8880
      OleObjectBlob   =   "frmNikeFootball.frx":15F942
      SourceDoc       =   "M:\CS130\AndrewEisingerVBproject\ESPN_Monday_Night_Football.mp3"
      TabIndex        =   7
      Top             =   6960
      Width           =   2895
   End
End
Attribute VB_Name = "frmNikeFootball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' NikeFootball
' Andrew Eisinger
' 3/16/09
'This program reads a file
'This program then based on the file can display and then sort by price and alphabetically
Dim FootballItems(1 To 250) As String, FootballCosts(1 To 250) As Single


Private Sub cmdGoBack_Click()
frmNike1.Show
frmNikeFootball.Hide
End Sub

Private Sub cmdGoBackHome_Click()
frmStoreHome.Show
frmNikeFootball.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReadData_Click()
picResults.Cls
Open App.Path & "\NikeFootball.txt" For Input As #1
CTR = 0
picResults.Print "Football Item"; Tab(40); "Football Cost"
picResults.Print "**************************************************************************"
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, FootballItems(CTR), FootballCosts(CTR)
    picResults.Print FootballItems(CTR); Tab(40); FormatCurrency(FootballCosts(CTR))
Loop
Close #1
End Sub


Private Sub cmdSortbyName_Click()
'This button sorts the arrays by the name of the products
Dim J As Single, TempFootballItems As String, TempFootballCosts As Single
    'Clears the previous results
    picResults.Cls
    
    'Prints headings so it is easier for the user to understand
    'the results.
    picResults.Print "Football Item"; Tab(40); "Football Cost"
    picResults.Print "**************************************************************************"
    
    
    'Code to sort the two parralel arrays by the first name of the product
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If FootballItems(Pos) > FootballItems(Pos + 1) Then
                TempFootballItems = FootballItems(Pos)
                FootballItems(Pos) = FootballItems(Pos + 1)
                FootballItems(Pos + 1) = TempFootballItems
                TempFootballCosts = FootballCosts(Pos)
                FootballCosts(Pos) = FootballCosts(Pos + 1)
                FootballCosts(Pos + 1) = TempFootballCosts
            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print FootballItems(J); Tab(40); FormatCurrency(FootballCosts(J))
    Next
End Sub

Private Sub cmdSortbyPrice_Click()
'This button sorts the arrays by the price of the products
Dim J As Single, TempFootballItems As String, TempFootballCosts As Single
    'Clears the previous results
    picResults.Cls
    
    'Prints headings so it is easier for the user to understand
    'the results.
    picResults.Print "Football Item"; Tab(40); "Football Cost"
    picResults.Print "**************************************************************************"
    
    
    'Code to sort the two parralel arrays by the price of the product
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If FootballCosts(Pos) > FootballCosts(Pos + 1) Then
                TempFootballCosts = FootballCosts(Pos)
                FootballCosts(Pos) = FootballCosts(Pos + 1)
                FootballCosts(Pos + 1) = TempFootballCosts
                TempFootballItems = FootballItems(Pos)
                FootballItems(Pos) = FootballItems(Pos + 1)
                FootballItems(Pos + 1) = TempFootballItems
            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print FootballItems(J); Tab(40); FormatCurrency(FootballCosts(J))
    Next
End Sub
