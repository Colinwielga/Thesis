VERSION 5.00
Begin VB.Form frmSlopestylePrelims 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Slopestyle Preliminaries"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Return to main page"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Find Average"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for year"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Results"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   3855
      Left            =   2400
      ScaleHeight     =   3795
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   240
      Width           =   6735
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H000000C0&
      Caption         =   "Load Data"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmSlopestylePrelims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'page allows user to load preliminary slopestyle scores and print results with the year
'and two runs. Allows the user to search for particular year's scores and will find the average score.
'Also contains a button to bring thuser to the home page.

'this function adds the scores of both runs and finds the average score. The increment
'is 2 because there are 2 runs involved. Then prints results.
Private Sub cmdCompute_Click()
    Inc = 0
    sum = 0
    For Pos = 1 To ctr
        Inc = Inc + 2
        sum = sum + Run1(Pos) + Run2(Pos)
    Next Pos
    picResults.Cls
    Avg = sum / Inc
    picResults.Print "Shaun White's average slopestyle preliminary score is: "; FormatNumber(Avg)

End Sub
'brings user to main page
Private Sub cmdExit_Click()
    frmShaunWhite.Show
    frmSlopestylePrelims.Hide
End Sub
'Loads data from text file into array
Private Sub cmdLoad_Click()
    Open App.Path & "\slopestylePrelim.txt" For Input As #1
    ctr = 0
    
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, Year(ctr), Run1(ctr), Run2(ctr)
    Loop
    Close #1

End Sub
'prints the data from the loaded text file, which is now an array
Private Sub cmdPrint_Click()
    picResults.Cls
    picResults.Print "Year", "Run 1", "Run 2"
    For Pos = 1 To ctr
        picResults.Print Year(Pos), Run1(Pos), Run2(Pos)
    Next Pos
End Sub
'user enters a year via input box and the year is searched for within the array.
'when the specific year is found it will print the results, otherwise if not found
'an error message is received
Private Sub cmdSearch_Click()
    SYear = InputBox("Enter a year from 2003-2007", "search")
    picResults.Cls
    Found = False
    Pos = 0
    Do While (Found = False And Pos < ctr)
        Pos = Pos + 1
        If Year(Pos) = SYear Then
            Found = True
        End If
    Loop
    
    If Found = True Then
        picResults.Print "Year"; Year(Pos); " Shaun scored: "; FormatNumber(Run1(Pos)); " and "; FormatNumber(Run2(Pos))
    Else
        MsgBox "Error: year not found!"
    End If
    
End Sub
