VERSION 5.00
Begin VB.Form frmWhereNow 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000C0C0&
      Caption         =   "Exit the Program"
      Height          =   1215
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   5175
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0000C0C0&
      Caption         =   "Go Back to the Main Menu"
      Height          =   1215
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   5175
   End
   Begin VB.CommandButton cmdWhereAreNow 
      BackColor       =   &H0000C0C0&
      Caption         =   "Search Where They Are Now!"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   4695
   End
   Begin VB.CommandButton cmdLoadData 
      BackColor       =   &H0000C0C0&
      Caption         =   "Load the Data"
      Height          =   1815
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   120
      Picture         =   "frmWhereNow.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   7155
      TabIndex        =   4
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmWhereNow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Heisman Trophy
'frmWhereNow
'Kevin Abbas
'2-21-10
'Objective of form -  To let the user search for a Heisman winner and output what that person is doing now.


Dim Year(1 To 1000) As Integer, Player(1 To 1000) As String, WhereNow(1 To 1000) As String, Ctr As Integer

Option Explicit

Private Sub cmdExit_Click() 'exit the program
    MsgBox ("Hope you enjoyed learning about the Heisman, have a nice day!")
    End
End Sub

Private Sub cmdLoadData_Click() 'load the data from the data file


    Open App.Path & "\WhereNow2.txt" For Input As #1
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Year(Ctr), Player(Ctr), WhereNow(Ctr)
    Loop
    MsgBox ("The Data has been Entered!")
    Close #1
    cmdLoadData.Enabled = False
    cmdWhereAreNow.Enabled = True
   
    
End Sub

Private Sub cmdMainMenu_Click() 'bring the user back to the main menu
    frmWelcome.Show
    frmHistory.Hide
    frmWinners.Hide
    frmWhereNow.Hide
End Sub

Private Sub cmdWhereAreNow_Click() 'search the data file and display where the player is now - if the user enters an invalid name display an error message
    Dim S As String, K As Integer, Found As Boolean
        S = InputBox("Please enter the first and last name of whom you would like to search")
        Found = False
        For K = 1 To Ctr
            If Player(K) = S Then
                MsgBox ("Sweet! " & Player(K) & " won the Heisman in " & Year(K) & ". " & WhereNow(K))
                Found = True
            End If
        Next K
        If Found = False Then
            MsgBox ("Sorry! Noone knows what " & S & " is up to!")
        End If
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Form_Load()
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
