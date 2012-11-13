VERSION 5.00
Begin VB.Form frmWinners 
   BackColor       =   &H8000000D&
   Caption         =   "Winners"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCollege 
      Caption         =   "Sort by College Name"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Go Back To Main Menu"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Sort by Total Won"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlphabet 
      Caption         =   "Sort Winner Name Alphabetically"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load and Display Winners"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000003&
      Height          =   6375
      Left            =   4440
      ScaleHeight     =   6315
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
End
Attribute VB_Name = "frmWinners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: CSB/SJU Jeopardy
'Form Name: frmWinners
'Authors: Emma Jaynes, Lindsay Havlik, Brooke Beyer
'Date Written: 11/02/08
'Objective: This form loads the top 20 college winners of all time and is able to sort them according to what button is pushed
'Comments: Buttons perform different array sorts: alphabetically, by total, and by school name.  There is also a button to take you back to main menu

Option Explicit
Dim Names(1 To 20) As String, College(1 To 20) As String, Total(1 To 20) As Single
Dim CTR As Integer, TempName As String, TempCollege As String, TempTotal As Single
Dim Pass As Integer, Pos As Integer, N As Integer

Private Sub cmdAlphabet_Click()
'sorts array alphabetically

picOutput.Cls

picOutput.Print "Name:"; Tab(25); "College:"; Tab(60); "Total:"
picOutput.Print "*****************************************************************************************"

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Names(Pos) > Names(Pos + 1) Then
            TempName = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = TempName
            TempCollege = College(Pos)
            College(Pos) = College(Pos + 1)
            College(Pos + 1) = TempCollege
            TempTotal = Total(Pos)
            Total(Pos) = Total(Pos + 1)
            Total(Pos + 1) = TempTotal
        End If
    Next Pos
Next Pass

For N = 1 To 20
    picOutput.Print Names(N); Tab(25); College(N); Tab(60); FormatCurrency(Total(N))
Next N

End Sub

Private Sub cmdCollege_Click()
'sorts array alphabetically by college name

picOutput.Cls

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If College(Pos) > College(Pos + 1) Then
            TempCollege = College(Pos)
            College(Pos) = College(Pos + 1)
            College(Pos + 1) = TempCollege
            TempTotal = Total(Pos)
            Total(Pos) = Total(Pos + 1)
            Total(Pos + 1) = TempTotal
            TempName = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = TempName
        End If
    Next Pos
Next Pass

picOutput.Print "Name:"; Tab(25); "College:"; Tab(60); "Total:"
picOutput.Print "******************************************************************************************"
        
For N = 1 To 20
    picOutput.Print Names(N); Tab(25); College(N); Tab(60); FormatCurrency(Total(N))
Next N
    
End Sub

Private Sub cmdLoad_Click()
'loads that data to multiple arrays and prints that data

picOutput.Cls

CTR = 0

Open App.Path & "\winners.txt" For Input As #1

picOutput.Print "Name:"; Tab(25); "College:"; Tab(60); "Total:"
picOutput.Print "*****************************************************************************************"

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Names(CTR), College(CTR), Total(CTR)
    picOutput.Print Names(CTR); Tab(25); College(CTR); Tab(60); FormatCurrency(Total(CTR))
Loop
Close #1

End Sub

Private Sub cmdMainMenu_Click()
'takes us back to main menu

frmMainMenu.Show
frmWinners.Hide

End Sub

Private Sub cmdTotal_Click()
'sorts array by total winnings

picOutput.Cls

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Total(Pos) < Total(Pos + 1) Then
            TempTotal = Total(Pos)
            Total(Pos) = Total(Pos + 1)
            Total(Pos + 1) = TempTotal
            TempName = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = TempName
            TempCollege = College(Pos)
            College(Pos) = College(Pos + 1)
            College(Pos + 1) = TempCollege
        End If
    Next Pos
Next Pass

picOutput.Print "Name:"; Tab(25); "College:"; Tab(60); "Total:"
picOutput.Print "****************************************************************************************"

For N = 1 To 20
    picOutput.Print Names(N); Tab(25); College(N); Tab(60); FormatCurrency(Total(N))
Next N
        
End Sub

