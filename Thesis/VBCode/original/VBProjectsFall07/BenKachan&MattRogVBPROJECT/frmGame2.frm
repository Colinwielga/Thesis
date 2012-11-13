VERSION 5.00
Begin VB.Form frmGame2 
   Caption         =   "France v. Germany"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   Picture         =   "frmGame2.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdRetrun 
      Caption         =   "Return to Main Menu"
      Height          =   1095
      Left            =   600
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.PictureBox picIns 
      Height          =   1935
      Left            =   3480
      ScaleHeight     =   1875
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton cmdCommence 
      Caption         =   "Begin Shootout"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdInst 
      Caption         =   "Instructions"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmGame2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCommence_Click()
    'This program allows the user to enter a number from a small range 1-5 designating their shot on goal
    'Through If statements and nested If statements we were able to deide whether the shot was a goal, whether a six shooter is necessary, and whether or not you won the shootout
    Dim first As Integer, sum As Integer, second As Integer
    Dim third As Integer, fourth As Integer, fifth As Integer, sixth As Integer
    first = InputBox("Your first shooter is Patrick Vieira. Where would you like him to shoot 1-5")
    sum = 0
    picIns.Cls
    If (first = 3 Or first = 4 Or first = 5) Then
        MsgBox "GGGGOOOOAAAAAALLLLL!!!!"
        sum = sum + 1
    End If
    If (first > 5 Or first < 1) Then
            MsgBox "You missed the net"
    End If
    If (first = 2 Or first = 1) Then
        MsgBox "What a great save by the keeper"
    End If
    second = InputBox("Your second shooter is Lilian Thuram. Where would you like him to shoot 1-5")
    
    If (second = 1 Or second = 2 Or second = 5) Then
        
        MsgBox "An unbelievable save by the goalie"
    End If
    If (second > 5 Or second < 1) Then
            MsgBox "You missed the net"
    End If
    If (second = 4 Or second = 3) Then
        MsgBox "He beat him right down the middle of the goal, GOAL!"
        sum = sum + 1
    End If
    third = InputBox("Your third shooter is Franck Ribery. Where would you like him to shoot 1-5")
    
    If (third = 1 Or third = 5) Then
        sum = sum + 1
        MsgBox "Too much technique from Ribery as he finesses it past the keeper"
    End If
    If (third = 4 Or third = 2) Then
        MsgBox "What a poor shot, save"
    End If
    If (third > 5 Or third < 1) Then
            MsgBox "You missed the net"
    End If
    If (third = 3) Then
        sum = sum + 1
        MsgBox "It looks to be a save, no wait the ball has trickled over the line, GOAL!"
    End If
    fourth = InputBox("Your fourth shooter is Thierry Henry. Where would you like him to shoot 1-5")
    
    If (fourth = 2 Or fourth = 4) Then
        
        MsgBox "He shot it right at the goalie, easy save"
    End If
    If (fourth > 5 Or fourth < 1) Then
            MsgBox "You missed the net"
    End If
    If (fourth = 1 Or fourth = 3 Or fourth = 5) Then
        MsgBox "Thierry Henry fires it into the corner of the net, GOAL!"
        sum = sum + 1
    End If
    fifth = InputBox("Your fifth shooter is the controversial Zinidine Zidane. Where would you like him to shoot 1-5")
    
    If (fifth = 3) Then
        MsgBox "He chipped the keeper from the penalty spot!! GOAL!"
        sum = sum + 1
    End If
    If (fifth = 1 Or fifth = 2 Or fifth = 4 Or fifth = 5) Then
        MsgBox "Zidane decided to head butt ref instead of shoot"
    End If
    If (fifth > 5 Or fifth < 1) Then
            MsgBox "Zidane decided to head butt ref instead of shoot"
    End If
    If (sum > 2) Then
        picIns.Print "Congratulations you won the Collegeville Open Shootout Championship"
        picIns.Print "You managed to score: " & sum; " goals "
    End If
    If (sum = 2) Then
        sixth = InputBox("its all tied up and its up to your sixth shooter Claude Makelele")
        If (sixth = 1 Or sixth = 5 Or sixth = 4) Then
            picIns.Print "He did it. Congratulations you won the Collegeville Open Shootout Championship"
            sum = sum + 1
            picIns.Print "You managed to score: " & sum; " goals "
        End If
        If (sixth = 2 Or sixth = 3) Then
            picIns.Print "You gave it your all but Germany was the better team today, you lose."
           picIns.Print "You managed to score: " & sum; " goals "
        End If
    If (sixth > 5 Or sixth < 1) Then
            picIns.Print "You gave it your all but Germany was the better team today, you lose."
           picIns.Print "You managed to score: " & sum; " goals "
    End If
    End If
    If (sum < 2) Then
        picIns.Print "You gave it your all but Germany was the better team today, you lose."
        picIns.Print "You managed to score: " & sum; " goals "
    End If
End Sub

Private Sub cmdInst_Click()
    'Gives instructions for the Shooout procedures
    picIns.Cls
    picIns.Print "The shootout is played as follows.  "
    picIns.Print "You will select a location for your shot by choosing "
    picIns.Print "a number 1-5 which determines where the player will shoot"
    picIns.Print "you score by guessing the correct location to shoot."
End Sub

Private Sub cmdRetrun_Click()
    'return to Main menu
    frmGame2.Hide
    frmHome.Show
End Sub
