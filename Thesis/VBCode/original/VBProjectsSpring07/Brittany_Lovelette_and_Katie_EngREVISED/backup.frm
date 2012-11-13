VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Option Explicit

Private Sub cmdAge_Click()
    Dim Currentdate As Integer
    Dim Birthdate As Integer
    Dim Age As Integer
    Dim Birthyear As Integer
    Dim Currentyear As Integer
    Currentyear = 2007
    
    picResults.Cls
    
    
    Birthyear = InputBox("What year were you born? (yyyy)", "Your Birth Year")
    Currentdate = InputBox("What is the date? (mmdd)", "Current Date")
    Birthdate = InputBox("What is your birthdate? (mmdd)", "Birthdate")
    Select Case Currentdate
        Case Is < Birthdate
            Age = (Currentyear - Birthyear) - 1
        Case Is > Birthdate
            Age = Currentyear - Birthyear
        Case Is = Birthdate
            Age = Currentyear - Birthyear
                picResults.Print "Today is your birthday!"; Tab(2); "Happy Birthday!"
        End Select
    
    picResults.Print "You are "; Age; " years old."
    
    
    
End Sub

Private Sub cmdBirthstone_Click()
    Dim Month As String
    Dim Birthstone As String
    Dim January As String
    Dim February As String
    Dim March As String
    Dim April As String
    Dim May As String
    Dim June As String
    Dim July As String
    Dim August As String
    Dim September As String
    Dim October As String
    Dim November As String
    Dim December As String
    
    picResults.Cls
    
    
    Month = InputBox("Enter your birth month (ex: January)", "Birth Month")
    
    
    picResults.Print "Your Birthstone is:"
    picResults.Print
    
    Select Case Month
        Case Is = "January"
            picResults.Print "Garnet"
        Case Is = "February"
            picResults.Print "Amethyst"
        Case Is = "March"
            picResults.Print "Bloodstone or Aquamarine"
        Case Is = "April"
            picResults.Print "Diamond"
        Case Is = "May"
            picResults.Print "Emerald"
        Case Is = "June"
            picResults.Print "Pearl or Moonstone"
        Case Is = "July"
            picResults.Print "Ruby"
        Case Is = "August"
            picResults.Print "Sardonyx or Peridot"
        Case Is = "September"
            picResults.Print "Sapphire or Lapis Lazuli"
        Case Is = "October"
            picResults.Print "Opal or Pink Tourmaline"
        Case Is = "November"
            picResults.Print "Topaz or Citrine"
        Case Is = "December"
            picResults.Print "Turquoise or Zircon"
        Case Else
            picResults.Print "You have entered an invalid month."
    End Select

End Sub

Private Sub cmdClear_Click()
    picResults.Cls
    
End Sub





Private Sub cmdHoroscope_Click()
    Dim Ctr As Integer
    Dim Pos As Integer
    Dim Found As Boolean
    Dim Dates(1 To 366) As String
    Dim Horoscopes(1 To 100) As String
    Dim Birthdate As String
    
    Ctr = 0
    
    Open App.Path & "\dateswithhoroscope.txt" For Input As #1
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Dates(Ctr), Horoscopes(Ctr)
    Loop
    Close #1
    
    picResults.Cls
    
    Birthdate = InputBox("Enter your birthdate (Format January X):", "Birthdate")
    Found = False
    Pos = 0
    
    Do While (Found = False And Pos < Ctr)
        Pos = Pos + 1
        If Dates(Pos) = Birthdate Then
            Found = True
        End If
    Loop
    
    If Found = True Then
        picResults.Print "Your Horoscope is:"
        picResults.Print "---------------------------------"
        picResults.Print
        picResults.Print Horoscopes(Pos)
    Else
        picResults.Print Birthdate; " is not valid."
    End If
End Sub

Private Sub cmdQuit_Click()
    MsgBox "Thanks to www.astrology.com for horoscopes!", , "References"
    
    
    
    End
End Sub

End Sub
