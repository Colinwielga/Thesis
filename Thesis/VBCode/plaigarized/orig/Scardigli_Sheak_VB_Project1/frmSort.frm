VERSION 5.00
Begin VB.Form frmSort 
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search For Your Favorite Pitchers Stats"
      Height          =   1095
      Left            =   5760
      TabIndex        =   8
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdM 
      Caption         =   "Go To List Page"
      Height          =   975
      Left            =   9000
      TabIndex        =   7
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort5 
      Caption         =   "Sort By Loses"
      Height          =   975
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   9000
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort4 
      Caption         =   "Sort By ERA"
      Height          =   975
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort3 
      Caption         =   "Sort By Strike Outs"
      Height          =   975
      Left            =   8760
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdSort2 
      Caption         =   "Sort By Wins"
      Height          =   975
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdSort1 
      Caption         =   "Sort By First Name"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6915
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   8505
      Left            =   0
      Picture         =   "frmSort.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10740
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Cy Young Award Winners Over the Last 30 Years
'Form Name: frmSort
'Author: Anthony and Cameron
'Date Written: February 13, 2010
'Objective: This form sorts the parts of the array in a few different ways and allows the user to search the array.


Private Sub cmdM_Click()
    frmList.Show 'return to the list page
    frmSort.Hide 'hide the sort page
    
    
End Sub

Private Sub cmdQuit_Click()
    End 'quit the program
End Sub



Private Sub cmdSearch_Click()
    Dim A As Integer, X As String, Found As Boolean 'variables used in this search
    
    Found = False
    X = InputBox("Enter the name of your favorite Cy Young Award winner.") 'use an input box to input name
    
    Do Until Found = True 'run search until it finds the name of the correct pitcher
    
        A = A + 1
        If X = Pitchers(A) Then
            Found = True
            MsgBox ("Your favorite pitcher won " & Wins(A) & " games the year that he won the Cy Young.") 'this prints after the pitcher was found
            
        End If
    Loop
    
    If Not Found Then 'display this when the inputted pitcher was not found
        MsgBox ("There were no pitchers that won the Cy Young with that name.")
    End If
        
    
            
    
    
    
        
End Sub

Private Sub cmdSort1_Click()
'sort all the Cy Young winners by first name

    Dim Pass As Integer, Temp As String, N As Integer, Temp2 As String, Temp3 As String, Temp4 As String
    picResults.Cls 'clear the results screen
    For Pass = 1 To 29 'sort the array alphabetically
        For N = 1 To 30 - Pass
            If Pitchers(N) > Pitchers(N + 1) Then
            Temp = Pitchers(N)
                Pitchers(N) = Pitchers(N + 1)
                Pitchers(N + 1) = Temp
                Temp2 = Wins(N)
                Wins(N) = Wins(N + 1)
                Wins(N + 1) = Temp2
                Temp3 = Losses(N)
                Losses(N) = Losses(N + 1)
                Losses(N + 1) = Temp3
                Temp3 = ERA(N)
                ERA(N) = ERA(N + 1)
                ERA(N + 1) = Temp3
                Temp4 = StrikeOuts(N)
                StrikeOuts(N) = StrikeOuts(N + 1)
                StrikeOuts(N + 1) = Temp4
            End If
        Next N
    Next Pass
    
    MsgBox ("Sort Complete")
    picResults.Print "Pitchers", "Wins", "Losses", "ERA", "Strike Outs"
    picResults.Print "*************************************************************************************"
    
    For N = 1 To 30
        picResults.Print Pitchers(N); Tab(20); Wins(N), Losses(N), ERA(N), StrikeOuts(N) 'print the array alphabetically
    Next N
End Sub

Private Sub cmdSort2_Click()
Dim Pass As Integer, Temp As String, N As Integer, Temp2 As String, Temp3 As String, Temp4 As String
picResults.Cls
    For Pass = 1 To 29
        For N = 1 To 30 - Pass
            If Wins(N) < Wins(N + 1) Then 'sort the array by wins from most to least
            Temp = Wins(N)
                Wins(N) = Wins(N + 1)
                Wins(N + 1) = Temp
                Temp2 = ERA(N)
                ERA(N) = ERA(N + 1)
                ERA(N + 1) = Temp2
                Temp3 = Losses(N)
                Losses(N) = Losses(N + 1)
                Losses(N + 1) = Temp3
                Temp3 = Pitchers(N)
                Pitchers(N) = Pitchers(N + 1)
                Pitchers(N + 1) = Temp3
                Temp4 = StrikeOuts(N)
                StrikeOuts(N) = StrikeOuts(N + 1)
                StrikeOuts(N + 1) = Temp4
            End If
        Next N
    Next Pass
    
    MsgBox ("Sort Complete")
    picResults.Print "Pitchers", "Wins", "Losses", "ERA", "Strike Outs"
    picResults.Print "***************************************************************************************"
    
    For N = 1 To 30
        picResults.Print Pitchers(N); Tab(20); Wins(N), Losses(N), ERA(N), StrikeOuts(N) 'print wins from most to least
    Next N
End Sub

Private Sub cmdSort3_Click()
    Dim Pass As Integer, Temp As String, N As Integer, Temp2 As String, Temp3 As String, Temp4 As String
    picResults.Cls
    For Pass = 1 To 29
        For N = 1 To 30 - Pass
            If StrikeOuts(N) < StrikeOuts(N + 1) Then 'sort by which pitcher had the most strike outs
            Temp = StrikeOuts(N)
                StrikeOuts(N) = StrikeOuts(N + 1)
                StrikeOuts(N + 1) = Temp
                Temp2 = Wins(N)
                Wins(N) = Wins(N + 1)
                Wins(N + 1) = Temp2
                Temp3 = Losses(N)
                Losses(N) = Losses(N + 1)
                Losses(N + 1) = Temp3
                Temp3 = Pitchers(N)
                Pitchers(N) = Pitchers(N + 1)
                Pitchers(N + 1) = Temp3
                Temp4 = ERA(N)
                ERA(N) = ERA(N + 1)
                ERA(N + 1) = Temp4
            End If
        Next N
    Next Pass
    
    MsgBox ("Sort Complete")
    picResults.Print "Pitchers", "Wins", "Losses", "ERA", "Strike Outs"
    picResults.Print "****************************************************************************************"
    
    For N = 1 To 30
        picResults.Print Pitchers(N); Tab(20); Wins(N), Losses(N), ERA(N), StrikeOuts(N) 'print the array
    Next N
End Sub

Private Sub cmdSort4_Click()
    Dim Pass As Integer, Temp As String, N As Integer, Temp2 As String, Temp3 As String, Temp4 As String
    picResults.Cls
    For Pass = 1 To 29
        For N = 1 To 30 - Pass
            If ERA(N) > ERA(N + 1) Then 'sort by the lowest ERA
            Temp = ERA(N)
                ERA(N) = ERA(N + 1)
                ERA(N + 1) = Temp
                Temp2 = Wins(N)
                Wins(N) = Wins(N + 1)
                Wins(N + 1) = Temp2
                Temp3 = Losses(N)
                Losses(N) = Losses(N + 1)
                Losses(N + 1) = Temp3
                Temp3 = Pitchers(N)
                Pitchers(N) = Pitchers(N + 1)
                Pitchers(N + 1) = Temp3
                Temp4 = StrikeOuts(N)
                StrikeOuts(N) = StrikeOuts(N + 1)
                StrikeOuts(N + 1) = Temp4
            End If
        Next N
    Next Pass
    
    MsgBox ("Sort Complete")
    picResults.Print "Pitchers", "Wins", "Losses", "ERA", "Strike Outs"
    picResults.Print "**********************************************************************************************"
    
    For N = 1 To 30
        picResults.Print Pitchers(N); Tab(20); Wins(N), Losses(N), ERA(N), StrikeOuts(N) 'print array
    Next N
End Sub

Private Sub cmdSort5_Click()
    Dim Pass As Integer, Temp As String, N As Integer, Temp2 As String, Temp3 As String, Temp4 As String
    picResults.Cls
    For Pass = 1 To 29
        For N = 1 To 30 - Pass
            If Losses(N) > Losses(N + 1) Then 'sort by the least amount of losses
            Temp = Losses(N)
                Losses(N) = Losses(N + 1)
                Losses(N + 1) = Temp
                Temp2 = Wins(N)
                Wins(N) = Wins(N + 1)
                Wins(N + 1) = Temp2
                Temp3 = ERA(N)
                ERA(N) = ERA(N + 1)
                ERA(N + 1) = Temp3
                Temp3 = Pitchers(N)
                Pitchers(N) = Pitchers(N + 1)
                Pitchers(N + 1) = Temp3
                Temp4 = StrikeOuts(N)
                StrikeOuts(N) = StrikeOuts(N + 1)
                StrikeOuts(N + 1) = Temp4
            End If
        Next N
    Next Pass
    
    MsgBox ("Sort Complete")
    picResults.Print "Pitchers", "Wins", "Losses", "ERA", "Strike Outs"
    picResults.Print "*********************************************************************************************"
    
    For N = 1 To 30
        picResults.Print Pitchers(N); Tab(20); Wins(N), Losses(N), ERA(N), StrikeOuts(N) 'print array
    Next N
End Sub
