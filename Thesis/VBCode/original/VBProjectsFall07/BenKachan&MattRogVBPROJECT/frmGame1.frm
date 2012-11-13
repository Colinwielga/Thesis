VERSION 5.00
Begin VB.Form frmGame1 
   Caption         =   "England v. Argentina"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   Picture         =   "frmGame1.frx":0000
   ScaleHeight     =   6645
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdRetrun 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.PictureBox picInstruct 
      Height          =   1935
      Left            =   3720
      ScaleHeight     =   1875
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton cmdInstructions 
      Caption         =   "Instructions"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin Shootout"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "frmGame1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBegin_Click()
    'This program allows the user to enter a number from a small range 1-5 designating their shot on goal
    'Through If statements and nested If statements we were able to deide whether the shot was a goal, whether a six shooter is necessary, and whether or not you won the shootout
    Dim first As Integer, sum As Integer, second As Integer
    Dim third As Integer, fourth As Integer, fifth As Integer, sixth As Integer
    first = InputBox("Your first shooter is John Terry. Where would you like him to shoot 1-5")
    sum = 0
    picInstruct.Cls
    If (first = 1 Or first = 4 Or first = 5) Then
        MsgBox "GGGGOOOOAAAAAALLLLL!!!!"
        sum = sum + 1
    End If
    If (first > 5 Or first < 1) Then
            MsgBox "You missed the net"
    End If
    If (first = 2 Or first = 3) Then
        MsgBox "What a great save by the keeper"
    End If
    second = InputBox("Your second shooter is Wayne Rooney. Where would you like him to shoot 1-5")
    
    If (second = 1 Or second = 4 Or second = 5) Then
        
        MsgBox "The goalie robs Rooney(save)"
    End If
    If (second > 5 Or second < 1) Then
            MsgBox "You missed the net"
    End If
    If (second = 2 Or second = 3) Then
        MsgBox "Rooney finds the back of the old onion bag!!!!(goal)"
        sum = sum + 1
    End If
    third = InputBox("Your third shooter is Stephen Gerrard. Where would you like him to shoot 1-5")
    
    If (third = 1 Or third = 5) Then
        sum = sum + 1
        MsgBox "Too much power from Gerrard as he shoots it past the keeper"
    End If
    If (third = 4) Then
        MsgBox "What a poor shot, save"
    End If
    If (third > 5 Or third < 1) Then
            MsgBox "You missed the net"
    End If
    If (third = 2 Or third = 3) Then
        sum = sum + 1
        MsgBox "It looks to be a save, no wait the ball has trickled over the line, GOAL!"
    End If
    fourth = InputBox("Your fourth shooter is Frank Lampard. Where would you like him to shoot 1-5")
    
    If (fourth = 2 Or fourth = 4) Then
        
        MsgBox "He shot it right at the goalie, easy save"
    End If
    If (fourth > 5 Or fourth < 1) Then
            MsgBox "You missed the net"
    End If
    If (fourth = 1 Or fourth = 3 Or fourth = 5) Then
        MsgBox "Frank tucks it into the corner of the net, GOAL!"
        sum = sum + 1
    End If
    fifth = InputBox("Your fifth shooter is David Beckham. Where would you like him to shoot 1-5")
    
    If (fifth = 3) Then
        MsgBox "What a powerful drive to beat the keeper"
        sum = sum + 1
    End If
    If (fifth = 1 Or fifth = 2 Or fifth = 4 Or fifth = 5) Then
        MsgBox "Beckham chokes again, save for the goalie"
    End If
    If (fifth > 5 Or fifth < 1) Then
            MsgBox "You missed the net, Beckham chokes again"
    End If
    If (sum > 2) Then
        picInstruct.Print "Congratulations you won the Collegeville Open Shootout Championship"
        picInstruct.Print "You managed to score: " & sum; " goals "
    End If
    If (sum = 2) Then
        sixth = InputBox("its all tied up and its up to your sixth shooter Michael Owen")
        If (sixth = 1 Or sixth = 5 Or sixth = 4) Then
            picInstruct.Print "He did it. Congratulations you won the Collegeville Open Shootout Championship"
            sum = sum + 1
            picInstruct.Print "You managed to score: " & sum; " goals "
        End If
        If (sixth = 2 Or sixth = 3) Then
            picInstruct.Print "You gave it your all but Argentina was the better team today, you lose."
           picInstruct.Print "You managed to score: " & sum; " goals "
        End If
    If (sixth > 5 Or sixth < 1) Then
            picInstruct.Print "You gave it your all but Argentina was the better team today, you lose."
           picInstruct.Print "You managed to score: " & sum; " goals "
    End If
    End If
    If (sum < 2) Then
        picInstruct.Print "You gave it your all but Argentina was the better team today, you lose."
        picInstruct.Print "You managed to score: " & sum; " goals "
    End If
End Sub


Private Sub cmdInstructions_Click()
    'Gives instructions for the Shooout procedures
    picInstruct.Cls
    picInstruct.Print "The shootout is played as follows.  "
    picInstruct.Print "You will select a location for your shot by choosing "
    picInstruct.Print "a number 1-5 which determines where the player will shoot"
    picInstruct.Print "you score by guessing the correct location to shoot."
End Sub

Private Sub cmdRetrun_Click()
    'return to Main menu
    frmGame1.Hide
    frmHome.Show
End Sub
