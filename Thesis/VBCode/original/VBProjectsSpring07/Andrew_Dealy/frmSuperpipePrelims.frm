VERSION 5.00
Begin VB.Form frmSuperpipePrelims 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Superpipe Preliminaries"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main page"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Compute Average"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for year"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Results"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   2775
      Left            =   2160
      ScaleHeight     =   2715
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   240
      Width           =   6735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Data"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmSuperpipePrelims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'user loads data for superpipe preliminary runs and prints data. Finds the average score,
'allows user to search years, and has a link to the home page.

'this function adds the scores of both runs and finds the average score. The increment
'is 2 because there are 2 runs involved. Then prints results.
Private Sub cmdAverage_Click()
    Inc = 0
    sum = 0
    
    For Pos = 1 To ctr
        Inc = Inc + 2
        sum = sum + Run1(Pos) + Run2(Pos)
    Next Pos
    picResults.Cls
    Avg = sum / Inc
    picResults.Print "Shaun White's average superpipe preliminary score is: "; FormatNumber(Avg)
End Sub
'Loads data from text file into array
Private Sub cmdLoad_Click()
    Open App.Path & "\superpipePrelim.txt" For Input As #1
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
    picResults.Print "Year", "Run 1", "Run2"
    For Pos = 1 To ctr
        picResults.Print Year(Pos), Run1(Pos), Run2(Pos)
    Next Pos
End Sub
'brings user to main page
Private Sub cmdReturn_Click()
    frmShaunWhite.Show
    frmSuperpipePrelims.Hide
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
