VERSION 5.00
Begin VB.Form frmNikeRunning 
   BackColor       =   &H80000007&
   Caption         =   "NikeRunning"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   Picture         =   "frmNikeRunning.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000C&
      Caption         =   "Quit"
      Height          =   1335
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdBacktoHome 
      BackColor       =   &H8000000C&
      Caption         =   "Back To Store Home"
      Height          =   1335
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000C&
      Caption         =   "Back to Nike Store"
      Height          =   1335
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortbyCost 
      BackColor       =   &H8000000C&
      Caption         =   "Sort by Cost"
      Height          =   1335
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortbyName 
      BackColor       =   &H8000000C&
      Caption         =   "Sort By Name of Item"
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   4935
      Left            =   6600
      ScaleHeight     =   4875
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H8000000C&
      Caption         =   "Read Data"
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   2295
   End
End
Attribute VB_Name = "frmNikeRunning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
' NikeRunning
' Andrew Eisinger
' 3/17/09
'This program reads a file
'This program then based on the file can display and then sort by price and alphabetically
Dim RunningItems(1 To 50) As String, RunningCosts(1 To 50) As Single


Private Sub cmdBack_Click()
frmNike1.Show
frmNikeRunning.Hide
End Sub

Private Sub cmdBacktoHome_Click()
frmStoreHome.Show
frmNikeRunning.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()
picResults.Cls
Open App.Path & "\NikeShoes.txt" For Input As #1
CTR = 0
picResults.Print "Running Item"; Tab(40); "Running Cost"
picResults.Print "**************************************************************************"
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, RunningItems(CTR), RunningCosts(CTR)
    picResults.Print RunningItems(CTR); Tab(40); FormatCurrency(RunningCosts(CTR))
Loop
Close #1
End Sub

Private Sub cmdSortbyCost_Click()
'This button sorts the arrays by the price of the products
Dim J As Single, TempRunningItems As String, TempRunningCosts As Single
    'Clears the previous results
    picResults.Cls
    
    'Prints headings so it is easier for the user to understand
    picResults.Print "Running Item"; Tab(40); "Running Cost"
    picResults.Print "**************************************************************************"
    
    
    'Code to sort the two parralel arrays by the price of the product
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If RunningCosts(Pos) > RunningCosts(Pos + 1) Then
                TempRunningCosts = RunningCosts(Pos)
                RunningCosts(Pos) = RunningCosts(Pos + 1)
                RunningCosts(Pos + 1) = TempRunningCosts
                TempRunningItems = RunningItems(Pos)
                RunningItems(Pos) = RunningItems(Pos + 1)
                RunningItems(Pos + 1) = TempRunningItems
            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print RunningItems(J); Tab(40); FormatCurrency(RunningCosts(J))
    Next

End Sub

Private Sub cmdSortbyName_Click()
'This button sorts the arrays by the first names of the products
Dim J As Single, TempRunningItems As String, TempRunningCosts As Single
    'Clears the previous results
    picResults.Cls
    
    'Prints headings so it is easier for the user to understand
    picResults.Print "Running Item"; Tab(40); "Running Cost"
    picResults.Print "**************************************************************************"
   
    
    'Code to sort the two parralel arrays by the first name of the product
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If RunningItems(Pos) > RunningItems(Pos + 1) Then
                TempRunningItems = RunningItems(Pos)
                RunningItems(Pos) = RunningItems(Pos + 1)
                RunningItems(Pos + 1) = TempRunningItems
                TempRunningCosts = RunningCosts(Pos)
                RunningCosts(Pos) = RunningCosts(Pos + 1)
                RunningCosts(Pos + 1) = TempRunningCosts
            End If
        Next Pos
    Next Pass
    For J = 1 To CTR
        
        'Prints the results
        picResults.Print RunningItems(J); Tab(40); FormatCurrency(RunningCosts(J))
    Next

End Sub

