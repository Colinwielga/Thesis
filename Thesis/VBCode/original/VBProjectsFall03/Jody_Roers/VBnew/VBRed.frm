VERSION 5.00
Begin VB.Form Debt 
   BackColor       =   &H80000007&
   Caption         =   "Debt"
   ClientHeight    =   8700
   ClientLeft      =   885
   ClientTop       =   855
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   13080
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   840
      Picture         =   "VBRed.frx":0000
      ScaleHeight     =   2775
      ScaleWidth      =   1575
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to the Main Menu"
      Height          =   1215
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   9120
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "Which Clubs Are In Debt?"
      Height          =   1215
      Left            =   4800
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   3480
      ScaleHeight     =   1995
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   840
      Width           =   6375
   End
   Begin VB.Label lblName 
      Caption         =   "Created by Jody Roers"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDebt 
      BackColor       =   &H000000FF&
      Caption         =   "DEBT"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Debt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Club Auditing Aid (VBProject.vbp)
'Form Name: Debt (VBRed.frm)
'Author: Jody Roers
'Date Written: 27 October 2003
'Purpose: to determine if any clubs are in debt on a specific date in October
'   and if so how much are they in debt.

Private Sub cmdMenu_Click()
Debt.Hide 'return to the main menu form
Menu.Show

End Sub

Private Sub cmdQuit_Click()
End 'quitprogram
End Sub

Private Sub cmdRed_Click()
Dim Number As Integer
Dim J As Integer
Dim Balance(1 To 31) As Single
picResults.Cls 'clear screen
Number = InputBox("Enter October Date", "Date") 'ask for october date user wants to withdraw money on
Do While Number > 31 Or Number < 1 'if number is not between 1 and 31(because of 31 days in October) give error message
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    Number = InputBox("Enter October Date, only enter the number of the day specified.", "Date") 'ask for october date again
Loop 'continue until user enters number between 1 and 31
Open Menu.Path & "VB 10-21-02\Commtxt.txt" For Input As #1 'open file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
If Balance(Number) < 0 Then 'check if club is in debt on date specified above
        picResults.Print "The Communication Club is in Debt."
        picResults.Print Tab(10); "Their Balance is "; FormatCurrency(Balance(Number))
        'if in debt tell user and give user balance also
    Else
        picResults.Print "The Communication Club is not in Debt." 'if not in debt just print that
End If
Close #1 'close file
Open Menu.Path & "VB 10-21-02\Demstxt.txt" For Input As #1 'open next club's file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
If Balance(Number) < 0 Then 'check if club is in debt on date specified above
        picResults.Print "The College Democrats Club is in Debt." 'if in debt tell user and give user balance also
        picResults.Print Tab(10); "Their Balance is "; FormatCurrency(Balance(Number))
    Else
        picResults.Print "The College Democrats Club is not in Debt." 'if not in debt just print that
End If
Close #1 'close file
Open Menu.Path & "VB 10-21-02\Fusiontxt.txt" For Input As #1 'open next club's file
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
If Balance(Number) < 0 Then 'check if club is in debt on date specified above
        picResults.Print "The Cultural Fusion Club is in Debt." 'if in debt tell user and give user balance also
        picResults.Print Tab(10); "Their Balance is "; FormatCurrency(Balance(Number))
    Else
        picResults.Print "The Cultural Fusion Club is not in Debt." 'if not in debt just print that
End If
Close #1 'close file
Open Menu.Path & "VB 10-21-02\Rugbytxt.txt" For Input As #1
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
If Balance(Number) < 0 Then 'check if club is in debt on date specified above
        picResults.Print "The Women's Rugby Club is in Debt." 'if in debt tell user and give user balance also
        picResults.Print Tab(10); "Their Balance is "; FormatCurrency(Balance(Number))
    Else
        picResults.Print "The Women's Rugby Club is not in Debt." 'if not in debt just print that
End If
Close #1 'close file
Open Menu.Path & "VB 10-21-02\Spanishtxt.txt" For Input As #1
For J = 1 To 31 'fill array
    Input #1, Balance(J)
Next J
If Balance(Number) < 0 Then 'check if club is in debt on date specified above
        picResults.Print "The Spanish Club is in Debt." 'if in debt tell user and give user balance also
        picResults.Print Tab(10); "Their Balance is "; FormatCurrency(Balance(Number))
    Else
        picResults.Print "The Spanish Club is not in Debt." 'if not in debt just print that
End If
Close #1 'close file
End Sub

