VERSION 5.00
Begin VB.Form frmRules 
   BackColor       =   &H0000C000&
   Caption         =   "Rules Of Beer Ball"
   ClientHeight    =   12690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15855
   LinkTopic       =   "Form2"
   ScaleHeight     =   12690
   ScaleWidth      =   15855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Back to Main"
      Height          =   2055
      Left            =   720
      TabIndex        =   2
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CommandButton cmdshowrules 
      Caption         =   "See Rules"
      Height          =   1935
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.PictureBox picresults 
      Height          =   6375
      Left            =   5160
      ScaleHeight     =   6315
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   240
      Width           =   10335
   End
   Begin VB.Image Image1 
      Height          =   5730
      Left            =   600
      Picture         =   "rules.frx":0000
      Top             =   6720
      Width           =   7650
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form shows the rules of the Beerball game in a picture box
'it is intended for educational purposes only and the author of this program does not condone underage or binge drinking



Private Sub cmdreturn_Click()
'returns user back to main menu form
frmRules.Hide
frmmain.Show

End Sub


Private Sub cmdshowrules_Click()
' this subroutine shows the user the rules of Beerball
    Dim rules(1 To 6) As String 'loads rules in and array
    Dim CTR As Integer, pos As Integer

    picresults.Cls
    
    
    Open App.Path & "\rules.txt" For Input As #1 'loads the list of rules

    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, rules(CTR) 'puts the list of rules into an array
    Loop

    Close #1 'closes the rules.txt file

    For pos = 1 To CTR
        picresults.Print rules(pos) 'displays the list of rules
        picresults.Print "                                                                                               "
    Next pos
End Sub
