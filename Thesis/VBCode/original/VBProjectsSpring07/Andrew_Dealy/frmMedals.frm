VERSION 5.00
Begin VB.Form frmMedals 
   BackColor       =   &H00400000&
   Caption         =   "Medal History"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main page"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoadSuper 
      Caption         =   "Load and Print Superpipe"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load and Print Slopestyle"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox picMedals 
      BackColor       =   &H00C0E0FF&
      Height          =   3855
      Left            =   2640
      ScaleHeight     =   3795
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmMedals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this page allows the user to view Shaun's medal history. The two buttons load the data
'from the two competitions and print the results with the year and type of medal received

'this function loads the data from the text file into an array and prints the results.
Private Sub cmdLoad_Click()
    Open App.Path & "\slopeMedals.txt" For Input As #1
    ctr = 0
    
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, Year(ctr), Medal(ctr)
    Loop
    Close #1
    
    picMedals.Cls
    picMedals.Print "Slopestyle:"
    picMedals.Print "Year", "Medal"
    
    For Pos = 1 To ctr
        picMedals.Print Year(Pos), Medal(Pos)
    Next Pos
End Sub
'this function loads the data from the text file into and array and prints the results
Private Sub cmdLoadSuper_Click()
    Open App.Path & "\superMedals.txt" For Input As #1
    Inc = 0
    
    Do Until EOF(1)
        Inc = Inc + 1
        Input #1, Year(Inc), Medal(Inc)
    Loop
    Close #1
    
    picMedals.Cls
    picMedals.Print "Superpipe:"
    picMedals.Print "Year", "Medal"
    
    For Pos = 1 To Inc
        picMedals.Print Year(Pos), Medal(Pos)
    Next Pos
End Sub
'brings user back to main page
Private Sub cmdReturn_Click()
    frmShaunWhite.Show
    frmMedals.Hide
End Sub

