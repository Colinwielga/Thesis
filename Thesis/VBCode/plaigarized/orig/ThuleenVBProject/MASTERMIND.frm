VERSION 5.00
Begin VB.Form MASTERMIND 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form2"
   ScaleHeight     =   9495
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   17
      Top             =   6960
      Width           =   2175
   End
   Begin VB.PictureBox picResults55 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   16
      Top             =   3840
      Width           =   5175
   End
   Begin VB.PictureBox picResults44 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   15
      Top             =   3120
      Width           =   5175
   End
   Begin VB.PictureBox picResults33 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   14
      Top             =   2400
      Width           =   5175
   End
   Begin VB.PictureBox picResults22 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   13
      Top             =   1680
      Width           =   5175
   End
   Begin VB.PictureBox picResults11 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   12
      Top             =   960
      Width           =   5175
   End
   Begin VB.CommandButton cmdGuess5 
      Caption         =   "Guess 5"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   11
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start!"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5760
      TabIndex        =   10
      Top             =   4680
      Width           =   3855
   End
   Begin VB.PictureBox picResults5 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   9
      Top             =   3840
      Width           =   5175
   End
   Begin VB.PictureBox picResults4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   8
      Top             =   3120
      Width           =   5175
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   7
      Top             =   2400
      Width           =   5175
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   6
      Top             =   1680
      Width           =   5175
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   960
      Width           =   5175
   End
   Begin VB.CommandButton cmdGuess4 
      Caption         =   "Guess 4"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdGuess3 
      Caption         =   "Guess 3"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdGuess2 
      Caption         =   "Guess 2"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdGuess1 
      Caption         =   "Guess 1"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "MASTERMIND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim num1 As Integer, num2 As Integer, num3 As Integer, num4 As Integer, Pos As Integer

Private Sub cmdGuess1_Click()
    Dim guess1 As Integer, guess2 As Integer, guess3 As Integer, guess4 As Integer, CTR As Integer
    
    guess1 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess1 < 0 Or guess1 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess2 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess2 < 0 Or guess2 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess3 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess3 < 0 Or guess3 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess4 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess4 < 0 Or guess4 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    CTR = 0
    
    Do Until CTR = 4
        CTR = CTR + 1
    Loop
    picResults1.Print guess1; guess2; guess3; guess4;
        If guess1 = num1 Then
            picResults11.Print num1; " X"; " X"; " X"
        End If
        
        If guess1 = num2 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num3 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num4 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " X"; " X"
        End If
        
        If guess2 = num2 Then
            picResults11.Print "X "; num2; " X"; " X"
        End If
        
        If guess2 = num1 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num3 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num4 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " X"; " X"
        End If
        
        If guess3 = num3 Then
            picResults11.Print "X"; " X"; num3; " X"
        End If
        
        If guess3 = num1 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num2 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num4 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " !"; " X"
        End If
        
        If guess4 = num4 Then
            picResults11.Print "X"; " X"; " X"; num4
        End If
        
        If guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num2 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 Then
            picResults11.Cls
            picResults11.Print num1; num2; " X"; " X"
        End If
        
        If guess1 = num1 And guess3 = num3 Then
            picResults11.Cls
            picResults11.Print num1; " X"; num3; " X"
        End If
        
        If guess1 = num1 And guess4 = num4 Then
            picResults11.Cls
            picResults11.Print num1; " X"; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 Then
            picResults11.Cls
            picResults11.Print "X "; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess4 = num4 Then
            picResults11.Cls
            picResults11.Print "X "; num2; " X "; num4
        End If
        
        If guess3 = num3 And guess4 = num4 Then
            picResults11.Cls
            picResults11.Print "X"; " X "; num3; num4
        End If
        
        If guess1 = num2 And guess2 = num1 Then
            picResults11.Cls
            picResults11.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num3 And guess2 = num4 Then
            picResults11.Cls
            picResults11.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num4 And guess2 = num3 Then
            picResults11.Cls
            picResults11.Print "!"; " !"; " X"; " X"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " !"; " X"
        End If
    
        If guess2 = num4 And guess3 = num1 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " !"; " X"
        End If
        
        If guess3 = num1 And guess4 = num2 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num2 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num4 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "X"; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num3 And guess4 = num2 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num4 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " X"; " !"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num4 And guess3 = num1 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " !"; " X"
        End If
        
        If guess1 = num2 And guess3 = num4 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num3 And guess3 = num2 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num4 And guess3 = num1 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " !"; " X"
        End If
        
        If guess2 = num1 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " X"; " !"
        End If
    
        If guess2 = num3 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " X"; " !"
        End If
        
        If guess2 = num4 And guess4 = num2 Then
            picResults11.Cls
            picResults11.Print "X"; " !"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 Then
            picResults11.Cls
            picResults11.Print num1; num2; " !"; " X"
        End If

        If guess1 = num1 And guess2 = num2 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print num1; num2; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print num1; num2; " !"; " !"
        End If
            
        If guess1 = num3 And guess2 = num2 And guess3 = num4 Then
            picResults11.Cls
            picResults11.Print "! "; num2; " !"; " X"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 Then
            picResults11.Cls
            picResults11.Print num1; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess3 = num4 Then
            picResults11.Cls
            picResults11.Print "X "; num2; " !"; " X"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess1 = num4 Then
            picResults11.Cls
            picResults11.Print "! "; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "X "; num2; num3; " !"
        End If
        
        If guess2 = num4 And guess3 = num3 Then
            picResults11.Cls
            picResults11.Print "X "; " !"; num3; " X"
        End If
        
        If guess1 = num1 And guess2 = num4 And guess3 = num3 Then
            picResults11.Cls
            picResults11.Print num1; " !"; num3; " X"
        End If
        
        If guess3 = num3 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "X "; "X"; num3; " !"
        End If
        
        If guess1 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; num3; num4
        End If
        
        If guess1 = num2 And guess3 = num2 And guess4 = num4 Then
            picResults11.Cls
            picResults11.Print "!"; " X"; " !"; num4
        End If
        
        If guess1 = num1 And guess3 = num4 And guess4 = num2 Then
            picResults11.Cls
            picResults11.Print num1; " X"; " !"; " !"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults11.Cls
            picResults11.Print "X"; num2; num3; num4
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 = num3 And guess4 = num4 Then
            picResults11.Cls
            picResults11.Print num1; " X"; num3; num4
        End If
        
        If guess1 = num1 And guess2 = num2 And guess4 = num4 Then
            picResults11.Cls
            picResults11.Print num1; num2; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 And guess1 = num1 Then
            picResults11.Cls
            picResults11.Print num1; num2; num3; " X"
        End If
        
        If guess1 <> num1 And guess2 = num4 And guess3 = num2 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 <> num1 And guess2 = num3 And guess3 = num2 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 <> num1 And guess2 = num1 And guess3 = num4 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 <> num1 And guess2 = num4 And guess3 = num1 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 <> num1 And guess2 = num4 And guess3 = num2 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 <> num1 And guess2 = num3 And guess3 = num1 And guess4 = num2 Then
            picResults11.Cls
            picResults11.Print "X "; " !"; " !"; " !"
        End If
        
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults11.Cls
            picResults11.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 = num4 And guess3 <> num3 And guess4 = num3 Then
            picResults11.Cls
            picResults11.Print "! "; " !"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "Congratulations! You are a true genius!", vbOKOnly, "YOU WIN!"
        End If
           
    
            
End Sub

Private Sub cmdGuess2_Click()
Dim guess1 As Integer, guess2 As Integer, guess3 As Integer, guess4 As Integer, CTR As Integer
    
    guess1 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess1 < 0 Or guess1 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess2 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess2 < 0 Or guess2 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess3 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess3 < 0 Or guess3 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess4 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess4 < 0 Or guess4 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    CTR = 0
    
    Do Until CTR = 4
        CTR = CTR + 1
    Loop
    picResults2.Print guess1; guess2; guess3; guess4;
        If guess1 = num1 Then
            picResults22.Print num1 & " X"; " X"; " X"
        End If
        
        If guess1 = num2 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num3 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num4 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " X"; " X"
        End If
        
        If guess2 = num2 Then
            picResults22.Print "X "; num2 & " X"; " X"
        End If
        
        If guess2 = num1 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num3 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num4 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " X"; " X"
        End If
        
        If guess3 = num3 Then
            picResults22.Print "X"; " X"; num3 & " X"
        End If
        
        If guess3 = num1 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num2 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num4 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " !"; " X"
        End If
        
        If guess4 = num4 Then
            picResults22.Print "X"; " X"; " X"; num4
        End If
        
        If guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num2 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 Then
            picResults22.Cls
            picResults22.Print num1; num2; " X"; " X"
        End If
        
        If guess1 = num1 And guess3 = num3 Then
            picResults22.Cls
            picResults22.Print num1; " X"; num3; " X"
        End If
        
        If guess1 = num1 And guess4 = num4 Then
            picResults22.Cls
            picResults22.Print num1; " X"; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 Then
            picResults22.Cls
            picResults22.Print "X "; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess4 = num4 Then
            picResults22.Cls
            picResults22.Print "X "; num2; " X "; num4
        End If
        
        If guess3 = num3 And guess4 = num4 Then
            picResults22.Cls
            picResults22.Print "X"; " X "; num3; num4
        End If
        
        If guess1 = num2 And guess2 = num1 Then
            picResults22.Cls
            picResults22.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num3 And guess2 = num4 Then
            picResults22.Cls
            picResults22.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num4 And guess2 = num3 Then
            picResults22.Cls
            picResults22.Print "!"; " !"; " X"; " X"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " !"; " X"
        End If
    
        If guess2 = num4 And guess3 = num1 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " !"; " X"
        End If
        
        If guess3 = num1 And guess4 = num2 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num2 And guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num4 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "X"; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num3 And guess4 = num2 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num4 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " X"; " !"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num4 And guess3 = num1 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " !"; " X"
        End If
        
        If guess1 = num2 And guess3 = num4 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num3 And guess3 = num2 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num4 And guess3 = num1 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " !"; " X"
        End If
        
        If guess2 = num1 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " X"; " !"
        End If
    
        If guess2 = num3 And guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " X"; " !"
        End If
        
        If guess2 = num4 And guess4 = num2 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 Then
            picResults22.Cls
            picResults22.Print num1; num2; " !"; " X"
        End If

        If guess1 = num1 And guess2 = num2 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print num1; num2; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print num1; num2; " !"; " !"
        End If
            
        If guess1 = num3 And guess2 = num2 And guess3 = num4 Then
            picResults22.Cls
            picResults22.Print "! "; num2; " !"; " X"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 Then
            picResults22.Cls
            picResults22.Print num1; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess3 = num4 Then
            picResults22.Cls
            picResults22.Print "X"; num2; " !"; " X"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "X "; num2; num3; " !"
        End If
           
        If guess2 = num4 And guess3 = num3 Then
            picResults22.Cls
            picResults22.Print "X"; " !"; num3; " X"
        End If
        
        If guess3 = num3 And guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "X"; "X"; num3; " !"
        End If
        
        If guess1 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; num3; num4
        End If
        
        If guess1 = num2 And guess3 = num2 And guess4 = num4 Then
            picResults22.Cls
            picResults22.Print "!"; " X"; " !"; num4
        End If
        
        If guess1 = num1 And guess3 = num4 And guess4 = num2 Then
            picResults22.Cls
            picResults22.Print num1; " X"; " !"; " !"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults22.Cls
            picResults22.Print "X"; num2; num3; num4
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 = num3 And guess4 = num4 Then
            picResults22.Cls
            picResults22.Print num1; " X"; num3; num4
        End If
        
        If guess1 = num1 And guess2 = num2 And guess4 = num4 Then
            picResults22.Cls
            picResults22.Print num1; num2; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 And guess1 = num1 Then
            picResults22.Cls
            picResults22.Print num1; num2; num3; " X"
        End If
        
        If guess1 = num1 And guess2 = num4 And guess3 = num3 Then
            picResults22.Cls
            picResults22.Print num1; " !"; num3; " X"
        End If
        
        If guess1 <> num1 And guess2 = num4 And guess3 = num2 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults22.Cls
            picResults22.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 = num4 And guess3 <> num3 And guess4 = num3 Then
            picResults22.Cls
            picResults22.Print "! "; " !"; " X"; " !"
        End If



        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "Congratulations! You are a true genius!", vbOKOnly, "YOU WIN!"
        End If
End Sub

Private Sub cmdGuess3_Click()
Dim guess1 As Integer, guess2 As Integer, guess3 As Integer, guess4 As Integer, CTR As Integer
    
    guess1 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess1 < 0 Or guess1 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess2 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess2 < 0 Or guess2 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess3 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess3 < 0 Or guess3 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess4 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess4 < 0 Or guess4 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    CTR = 0
    
    Do Until CTR = 4
        CTR = CTR + 1
    Loop
    picResults3.Print guess1; guess2; guess3; guess4;
        If guess1 = num1 Then
            picResults33.Print num1 & " X"; " X"; " X"
        End If
        
        If guess1 = num2 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num3 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num4 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " X"; " X"
        End If
        
        If guess2 = num2 Then
            picResults33.Print "X "; num2 & " X"; " X"
        End If
        
        If guess2 = num1 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num3 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num4 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " X"; " X"
        End If
        
        If guess3 = num3 Then
            picResults33.Print "X"; " X"; num3 & " X"
        End If
        
        If guess3 = num1 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num2 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num4 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " !"; " X"
        End If
        
        If guess4 = num4 Then
            picResults33.Print "X"; " X"; " X"; num4
        End If
        
        If guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num2 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 Then
            picResults33.Cls
            picResults33.Print num1; num2; " X"; " X"
        End If
        
        If guess1 = num1 And guess3 = num3 Then
            picResults33.Cls
            picResults33.Print num1; " X"; num3; " X"
        End If
        
        If guess1 = num1 And guess4 = num4 Then
            picResults33.Cls
            picResults33.Print num1; " X"; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 Then
            picResults33.Cls
            picResults33.Print "X "; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess4 = num4 Then
            picResults33.Cls
            picResults33.Print "X "; num2; " X "; num4
        End If
        
        If guess3 = num3 And guess4 = num4 Then
            picResults33.Cls
            picResults33.Print "X"; " X "; num3; num4
        End If
        
        If guess1 = num2 And guess2 = num1 Then
            picResults33.Cls
            picResults33.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num3 And guess2 = num4 Then
            picResults33.Cls
            picResults33.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num4 And guess2 = num3 Then
            picResults33.Cls
            picResults33.Print "!"; " !"; " X"; " X"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " !"; " X"
        End If
    
        If guess2 = num4 And guess3 = num1 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " !"; " X"
        End If
        
        If guess3 = num1 And guess4 = num2 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num2 And guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num4 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "X"; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num3 And guess4 = num2 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num4 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " X"; " !"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num4 And guess3 = num1 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " !"; " X"
        End If
        
        If guess1 = num2 And guess3 = num4 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num3 And guess3 = num2 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num4 And guess3 = num1 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " !"; " X"
        End If
        
        If guess2 = num1 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " X"; " !"
        End If
    
        If guess2 = num3 And guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " X"; " !"
        End If
        
        If guess2 = num4 And guess4 = num2 Then
            picResults33.Cls
            picResults33.Print "X"; " !"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 Then
            picResults33.Cls
            picResults33.Print num1; num2; " !"; " X"
        End If

        If guess1 = num1 And guess2 = num2 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print num1; num2; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print num1; num2; " !"; " !"
        End If
            
        If guess1 = num3 And guess2 = num2 And guess3 = num4 Then
            picResults33.Cls
            picResults33.Print "! "; num2; " !"; " X"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 Then
            picResults33.Cls
            picResults33.Print num1; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess3 = num4 Then
            picResults33.Cls
            picResults33.Print "X "; num2; " !"; " X"
        End If
          
        If guess2 = num2 And guess3 = num3 And guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "X "; num2; num3; " !"
        End If
        
        If guess2 = num4 And guess3 = num3 Then
            picResults33.Cls
            picResults33.Print "X "; " !"; num3; " X"
        End If
        
        If guess3 = num3 And guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "X "; "X"; num3; " !"
        End If
        
        If guess1 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; num3; num4
        End If
        
        If guess1 = num2 And guess3 = num2 And guess4 = num4 Then
            picResults33.Cls
            picResults33.Print "!"; " X"; " !"; num4
        End If
        
        If guess1 = num1 And guess3 = num4 And guess4 = num2 Then
            picResults33.Cls
            picResults33.Print num1; " X"; " !"; " !"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults33.Cls
            picResults33.Print "X"; num2; num3; num4
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 = num3 And guess4 = num4 Then
            picResults33.Cls
            picResults33.Print num1; " X"; num3; num4
        End If
        
        If guess1 = num1 And guess2 = num2 And guess4 = num4 Then
            picResults33.Cls
            picResults33.Print num1; num2; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 And guess1 = num1 Then
            picResults33.Cls
            picResults33.Print num1; num2; num3; " X"
        End If
        
        If guess1 = num1 And guess2 = num4 And guess3 = num3 Then
            picResults33.Cls
            picResults33.Print num1; " !"; num3; " X"
        End If
        
        If guess1 <> num1 And guess2 = num4 And guess3 = num2 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults33.Cls
            picResults33.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 = num4 And guess3 <> num3 And guess4 = num3 Then
            picResults33.Cls
            picResults33.Print "! "; " !"; " X"; " !"
        End If



        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "Congratulations! You are a true genius!", vbOKOnly, "YOU WIN!"
        End If
End Sub

Private Sub cmdGuess4_Click()
Dim guess1 As Integer, guess2 As Integer, guess3 As Integer, guess4 As Integer, CTR As Integer

    
    guess1 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess1 < 0 Or guess1 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess2 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess2 < 0 Or guess2 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess3 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess3 < 0 Or guess3 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess4 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess4 < 0 Or guess4 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    CTR = 0
    
    Do Until CTR = 4
        CTR = CTR + 1
    Loop
    picResults4.Print guess1; guess2; guess3; guess4;
        If guess1 = num1 Then
            picResults44.Print num1 & " X"; " X"; " X"
        End If
        
        If guess1 = num2 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num3 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num4 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " X"; " X"
        End If
        
        If guess2 = num2 Then
            picResults44.Print "X "; num2 & " X"; " X"
        End If
        
        If guess2 = num1 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num3 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num4 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " X"; " X"
        End If
        
        If guess3 = num3 Then
            picResults44.Print "X"; " X"; num3 & " X"
        End If
        
        If guess3 = num1 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num2 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num4 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " !"; " X"
        End If
        
        If guess4 = num4 Then
            picResults44.Print "X"; " X"; " X"; num4
        End If
        
        If guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num2 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 Then
            picResults44.Cls
            picResults44.Print num1; num2; " X"; " X"
        End If
        
        If guess1 = num1 And guess3 = num3 Then
            picResults44.Cls
            picResults44.Print num1; " X"; num3; " X"
        End If
        
        If guess1 = num1 And guess4 = num4 Then
            picResults44.Cls
            picResults44.Print num1; " X"; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 Then
            picResults44.Cls
            picResults44.Print "X "; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess4 = num4 Then
            picResults44.Cls
            picResults44.Print "X "; num2; " X "; num4
        End If
        
        If guess3 = num3 And guess4 = num4 Then
            picResults44.Cls
            picResults44.Print "X"; " X "; num3; num4
        End If
        
        If guess1 = num2 And guess2 = num1 Then
            picResults44.Cls
            picResults44.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num3 And guess2 = num4 Then
            picResults44.Cls
            picResults44.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num4 And guess2 = num3 Then
            picResults44.Cls
            picResults44.Print "!"; " !"; " X"; " X"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " !"; " X"
        End If
    
        If guess2 = num4 And guess3 = num1 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " !"; " X"
        End If
        
        If guess3 = num1 And guess4 = num2 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num2 And guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num4 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "X"; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num3 And guess4 = num2 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num4 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " X"; " !"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num4 And guess3 = num1 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " !"; " X"
        End If
        
        If guess1 = num2 And guess3 = num4 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num3 And guess3 = num2 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num4 And guess3 = num1 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " !"; " X"
        End If
        
        If guess2 = num1 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " X"; " !"
        End If
    
        If guess2 = num3 And guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " X"; " !"
        End If
        
        If guess2 = num4 And guess4 = num2 Then
            picResults44.Cls
            picResults44.Print "X"; " !"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 Then
            picResults44.Cls
            picResults44.Print num1; num2; " !"; " X"
        End If

        If guess1 = num1 And guess2 = num2 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print num1; num2; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print num1; num2; " !"; " !"
        End If
            
        If guess1 = num3 And guess2 = num2 And guess3 = num4 Then
            picResults44.Cls
            picResults44.Print "! "; num2; " !"; " X"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 Then
            picResults44.Cls
            picResults44.Print num1; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess3 = num4 Then
            picResults44.Cls
            picResults44.Print "X "; num2; " !"; " X"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "X "; num2; num3; " !"
        End If
        
        If guess2 = num4 And guess3 = num3 Then
            picResults44.Cls
            picResults44.Print "X "; " !"; num3; " X"
        End If
        
        If guess3 = num3 And guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "X "; "X"; num3; " !"
        End If
        
        If guess1 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; num3; num4
        End If
        
        If guess1 = num2 And guess3 = num2 And guess4 = num4 Then
            picResults44.Cls
            picResults44.Print "!"; " X"; " !"; num4
        End If
        
        If guess1 = num1 And guess3 = num4 And guess4 = num2 Then
            picResults44.Cls
            picResults44.Print num1; " X"; " !"; " !"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults44.Cls
            picResults44.Print "X"; num2; num3; num4
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 = num3 And guess4 = num4 Then
            picResults44.Cls
            picResults44.Print num1; " X"; num3; num4
        End If
        
        If guess1 = num1 And guess2 = num2 And guess4 = num4 Then
            picResults44.Cls
            picResults44.Print num1; num2; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 And guess1 = num1 Then
            picResults44.Cls
            picResults44.Print num1; num2; num3; " X"
        End If
        
        If guess1 = num1 And guess2 = num4 And guess3 = num3 Then
            picResults44.Cls
            picResults44.Print num1; " !"; num3; " X"
        End If
        
        If guess1 <> num1 And guess2 = num4 And guess3 = num2 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults44.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults44.Cls
            picResults44.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 = num4 And guess3 <> num3 And guess4 = num3 Then
            picResults44.Cls
            picResults44.Print "! "; " !"; " X"; " !"
        End If



        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "Congratulations! You are a true genius!", vbOKOnly, "YOU WIN!"
        End If
End Sub

Private Sub cmdGuess5_Click()
Dim guess1 As Integer, guess2 As Integer, guess3 As Integer, guess4 As Integer, CTR As Integer

    
    guess1 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess1 < 0 Or guess1 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess2 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess2 < 0 Or guess2 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess3 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess3 < 0 Or guess3 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    guess4 = InputBox("Please enter a number 1-9", "Guess 1")
        If guess4 < 0 Or guess4 > 9 Then
            MsgBox "Please a number between 0-9", vbOKOnly, "Come On!"
        End If
    CTR = 0
    
    Do Until CTR = 4
        CTR = CTR + 1
    Loop
    picResults5.Print guess1; guess2; guess3; guess4;
        If guess1 = num1 Then
            picResults55.Print num1 & " X"; " X"; " X"
        End If
        
        If guess1 = num2 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num3 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " X"; " X"
        End If
        
        If guess1 = num4 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " X"; " X"
        End If
        
        If guess2 = num2 Then
            picResults55.Print "X "; num2 & " X"; " X"
        End If
        
        If guess2 = num1 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num3 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " X"; " X"
        End If
        
        If guess2 = num4 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " X"; " X"
        End If
        
        If guess3 = num3 Then
            picResults55.Print "X"; " X"; num3 & " X"
        End If
        
        If guess3 = num1 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num2 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " !"; " X"
        End If
        
        If guess3 = num4 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " !"; " X"
        End If
        
        If guess4 = num4 Then
            picResults55.Print "X"; " X"; " X"; num4
        End If
        
        If guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num2 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " X"; " !"
        End If
        
        If guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 Then
            picResults55.Cls
            picResults55.Print num1; num2; " X"; " X"
        End If
        
        If guess1 = num1 And guess3 = num3 Then
            picResults55.Cls
            picResults55.Print num1; " X"; num3; " X"
        End If
        
        If guess1 = num1 And guess4 = num4 Then
            picResults55.Cls
            picResults55.Print num1; " X"; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 Then
            picResults55.Cls
            picResults55.Print "X "; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess4 = num4 Then
            picResults55.Cls
            picResults55.Print "X "; num2; " X "; num4
        End If
        
        If guess3 = num3 And guess4 = num4 Then
            picResults55.Cls
            picResults55.Print "X"; " X "; num3; num4
        End If
        
        If guess1 = num2 And guess2 = num1 Then
            picResults55.Cls
            picResults55.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num3 And guess2 = num4 Then
            picResults55.Cls
            picResults55.Print "!"; " !"; " X"; " X"
        End If
        
        If guess1 = num4 And guess2 = num3 Then
            picResults55.Cls
            picResults55.Print "!"; " !"; " X"; " X"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " !"; " X"
        End If
    
        If guess2 = num4 And guess3 = num1 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " !"; " X"
        End If
        
        If guess3 = num1 And guess4 = num2 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num2 And guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " !"; " !"
        End If
        
        If guess3 = num4 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "X"; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num3 And guess4 = num2 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " X"; " !"
        End If
        
        If guess1 = num4 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " X"; " !"
        End If
        
        If guess2 = num1 And guess3 = num2 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num4 And guess3 = num1 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " !"; " X"
        End If
        
        If guess2 = num3 And guess3 = num4 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " !"; " X"
        End If
        
        If guess1 = num2 And guess3 = num4 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num3 And guess3 = num2 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " !"; " X"
        End If
        
        If guess1 = num4 And guess3 = num1 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " !"; " X"
        End If
        
        If guess2 = num1 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " X"; " !"
        End If
    
        If guess2 = num3 And guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " X"; " !"
        End If
        
        If guess2 = num4 And guess4 = num2 Then
            picResults55.Cls
            picResults55.Print "X"; " !"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 Then
            picResults55.Cls
            picResults55.Print num1; num2; " !"; " X"
        End If

        If guess1 = num1 And guess2 = num2 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print num1; num2; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num4 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print num1; num2; " !"; " !"
        End If
            
        If guess1 = num3 And guess2 = num2 And guess3 = num4 Then
            picResults55.Cls
            picResults55.Print "! "; num2; " !"; " X"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 Then
            picResults55.Cls
            picResults55.Print num1; num2; num3; " X"
        End If
        
        If guess2 = num2 And guess3 = num4 Then
            picResults55.Cls
            picResults55.Print "X "; num2; " !"; " X"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "X "; num2; num3; " !"
        End If
        
        If guess2 = num4 And guess3 = num3 Then
            picResults55.Cls
            picResults55.Print "X "; " !"; num3; " X"
        End If
        
        If guess3 = num3 And guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "X "; "X"; num3; " !"
        End If
        
        If guess1 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; num3; num4
        End If
        
        If guess1 = num2 And guess3 = num2 And guess4 = num4 Then
            picResults55.Cls
            picResults55.Print "!"; " X"; " !"; num4
        End If
        
        If guess1 = num1 And guess3 = num4 And guess4 = num2 Then
            picResults55.Cls
            picResults55.Print num1; " X"; " !"; " !"
        End If
        
        If guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults55.Cls
            picResults55.Print "X"; num2; num3; num4
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 = num3 And guess4 = num4 Then
            picResults55.Cls
            picResults55.Print num1; " X"; num3; num4
        End If
        
        If guess1 = num1 And guess2 = num2 And guess4 = num4 Then
            picResults55.Cls
            picResults55.Print num1; num2; " X"; num4
        End If
        
        If guess2 = num2 And guess3 = num3 And guess1 = num1 Then
            picResults55.Cls
            picResults55.Print num1; num2; num3; " X"
        End If
        
        If guess1 = num1 And guess2 = num4 And guess3 = num3 Then
            picResults55.Cls
            picResults55.Print num1; " !"; num3; " X"
        End If
        
        If guess1 <> num1 And guess2 = num4 And guess3 = num2 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "X "; " !"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 <> num2 And guess3 = num4 And guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num2 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num4 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num1 And guess4 = num2 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num3 And guess2 <> num2 And guess3 = num2 And guess4 = num1 Then
            picResults55.Cls
            picResults55.Print "! "; " X"; " !"; " !"
        End If
        
        If guess1 = num2 And guess2 = num4 And guess3 <> num3 And guess4 = num3 Then
            picResults55.Cls
            picResults55.Print "! "; " !"; " X"; " !"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 = num3 And guess4 <> num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 <> num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 = num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 <> num1 And guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 <> num1 And guess2 = num2 And guess3 = num3 And guess4 <> num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 = num3 And guess4 <> num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 = num1 And guess2 = num2 And guess3 <> num3 And guess4 <> num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 <> num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 <> num1 And guess2 = num2 And guess3 <> num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 <> num1 And guess2 <> num2 And guess3 = num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 <> num1 And guess2 <> num2 And guess3 <> num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 <> num1 And guess2 <> num2 And guess3 = num3 And guess4 <> num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 <> num1 And guess2 = num2 And guess3 <> num3 And guess4 <> num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If
        
        If guess1 = num1 And guess2 <> num2 And guess3 <> num3 And guess4 <> num4 Then
            picResults.Visible = True
            MsgBox "You Failed!", vbOKOnly, "LOSER!"
        End If


        If guess1 = num1 And guess2 = num2 And guess3 = num3 And guess4 = num4 Then
            picResults.Visible = True
            MsgBox "Congratulations! You are a true genius!", vbOKOnly, "YOU WIN!"
        End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStart_Click()
    Dim Pos As Integer
    picResults.Cls
    picResults1.Cls
    picResults11.Cls
    picResults2.Cls
    picResults22.Cls
    picResults3.Cls
    picResults33.Cls
    picResults4.Cls
    picResults44.Cls
    picResults5.Cls
    picResults55.Cls
    
    For Pos = 1 To 9
    
        Randomize
        num1 = Int((9 * Rnd) + 1)
        
        
        Randomize
        num2 = Int((9 * Rnd) + 1)
            If num2 = num1 Then
               num2 = Int((9 * Rnd) + 1)
            End If
            
        
        Randomize
        num3 = Int((9 * Rnd) + 1)
            If num2 = num3 Then
                num3 = Int((9 * Rnd) + 1)
            End If
            
            If num1 = num3 Then
                num3 = Int((9 * Rnd) + 1)
            End If
        
        
        Randomize
        num4 = Int(9 * Rnd)
            If num3 = num4 Then
                num4 = Int(9 * Rnd)
            End If
            
            If num2 = num4 Then
                num4 = Int(9 * Rnd)
            End If
            
            If num1 = num4 Then
                num4 = Int(9 * Rnd)
            End If
    Next Pos
    
    picResults.Print num1; num2; num3; num4
    picResults.Visible = False
End Sub

