VERSION 5.00
Begin VB.Form InTheEnd 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show me what I did!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "I want to try again!!!"
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   4320
      Width           =   1815
   End
   Begin VB.PictureBox picResultsEnd 
      Height          =   3375
      Left            =   360
      ScaleHeight     =   3315
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   360
      Width           =   8415
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "InTheEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project Name: Date Chooser (Wurm, Stacy - VB Project)
' Form Name: Dinner (StacyWurm5.frm)
' Author: Stacy Wurm
' Date Written: Wednesday, March 10th, 2004
' Purpose of this Form: ' To show the end results of the program
                        ' Lists what happened
                        ' Tells how much was spent and if user stayed within budget
                        
Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRestart_Click()
' go back to first choosing form
picResultsEnd.Cls
GiftToGive.Show
InTheEnd.Hide

' Reset the total values back to 0
CTR = 0
TotalCost = 0
MovieCost = 0
End Sub


Private Sub cmdShow_Click()
' Displays all the results for the whole program
picResultsEnd.Print UserName; ", here is what happened on your date!!"
picResultsEnd.Print
picResultsEnd.Print "You spent a total of "; FormatCurrency(TotalCost)
picResultsEnd.Print
Select Case TotalCost
    Case Is > Budget
        picResultsEnd.Print "Oops!!  You went over your budget!!"
        picResultsEnd.Print "You may want to start over and be more careful this time!!"
    Case Is < Budget
        picResultsEnd.Print "Yeah!!  You were able to stay within your budget on the date!!"
    Case Else
        picResultsEnd.Print "uh-oh"
End Select
picResultsEnd.Print
picResultsEnd.Print "You gave your date "; Decision1; " as a gift."
picResultsEnd.Print "You went to "; Decision2; " on your date."
    If Decision2 = "a movie" Then
        picResultsEnd.Print "If you chose to have concessions at the movie you spent "; FormatCurrency(MovieCost); " on them."
    End If
picResultsEnd.Print "Then you had dinner at "; Decision3; "."
picResultsEnd.Print "The date was ended by "; Decision4; "."
picResultsEnd.Print
picResultsEnd.Print "Hope you have a great time!!"
End Sub
