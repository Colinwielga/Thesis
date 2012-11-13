VERSION 5.00
Begin VB.Form frmThird 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   Picture         =   "frmThird.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdResource 
      Caption         =   "Game Theory Webpage"
      Height          =   615
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdStrategy 
      Caption         =   "Winning Strategy"
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   4080
      Picture         =   "frmThird.frx":FB34
      ScaleHeight     =   4395
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   960
      Width           =   5295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picResult 
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4395
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "frmThird"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Chen,Guo,Shi,Tian_Project1
'Form Name: frmFirst
'Author: Chen, Zhongjie
        'Guo, Zhishan
        'Shi, Yimei
        'Tian, Yukun
'Date Written: Oct. 24 to Oct. 26
'Objective: This form allows the user to play the Game with the computer
            'It also contains the most complicated algorithm
Option Explicit
Dim Objects(1 To 4) As Integer
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim SW_SHOW As Boolean, SW_NORMAL As Boolean
Private Sub cmdContinue_Click()
Dim Pile(1 To 4, 1 To 4) As Integer 'these 4 by 4 matrices will be used in the binary operation
Dim ObjectsAnalysis(1 To 4) As Integer
Dim i As Integer
Dim j As Integer
Dim M As Integer
Dim K As Integer
Dim Subtract As Integer
Dim P As Integer
Dim Q As Integer
Dim SubFrom As Integer
Dim NumSub As Integer
Dim Sum As Integer
picResult.Cls
Subtract = 0
If Objects(1) = 0 And Objects(2) = 0 And Objects(3) = 0 And Objects(4) = 0 Then 'this is the criterian for the user to win
    MsgBox "You Win! You are so smart!.", , "You win!"
Else
For i = 1 To 4
    For j = 1 To 4
        Pile(i, j) = 0
    Next j
Next i

'Binary operation starts here
For M = 1 To 4
    ObjectsAnalysis(M) = Objects(M)
Next M

For i = 1 To 4
    If ObjectsAnalysis(i) > 7 Then
        Pile(i, 1) = 1
        ObjectsAnalysis(i) = ObjectsAnalysis(i) - 8
    End If
    If ObjectsAnalysis(i) > 3 Then
        Pile(i, 2) = 1
        ObjectsAnalysis(i) = ObjectsAnalysis(i) - 4
    End If
    If ObjectsAnalysis(i) > 1 Then
        Pile(i, 3) = 1
        ObjectsAnalysis(i) = ObjectsAnalysis(i) - 2
    End If
    If ObjectsAnalysis(i) = 1 Then
        Pile(i, 4) = 1
    End If
Next i
'For i = 1 To 4
    'picResult.Print Pile(i, 1); Pile(i, 2); Pile(i, 3); Pile(i, 4)
'Next i
'If one wants to see the result of the binary operation, she/he can just include the two lines above into the code

'Find if there is any odd column
'Return to an even position

K = 1
For j = 1 To 4
    If Round((Pile(1, j) + Pile(2, j) + Pile(3, j) + Pile(4, j)) / 2) <> (Pile(1, j) + Pile(2, j) + Pile(3, j) + Pile(4, j)) / 2 Then 'this allows the computer to see if the column is odd
        K = 0
    End If
Next j
'If already in an even position, the computer has no smart move to make, so just take one object from the first non-zero pile
If K = 1 Then
    If Objects(1) > 0 Then
        Objects(1) = Objects(1) - 1
        MsgBox "I choose to remove one object from pile one", , "my choice"
    Else
        If Objects(2) > 0 Then
            Objects(2) = Objects(2) - 1
            MsgBox "I choose to remove one object from pile two", , "my choice"
        Else
            If Objects(3) > 0 Then
                Objects(3) = Objects(3) - 1
                MsgBox "I choose to remove one object from pile three", , "my choice"
            Else
                If Objects(4) > 0 Then
                MsgBox "I choose to remove one object from pile four", , "my choice"
                End If
            End If
        End If
    End If
    
Else 'start from here is the code when the original position is odd, which means the computer has a winning strategy
        'analyze one column after another, the code are similar for each one.
    Sum = Objects(1) + Objects(2) + Objects(3) + Objects(4)
    If Sum = Objects(1) Or Sum = Objects(2) Or Sum = Objects(3) Or Sum = Objects(4) Then
        If Objects(1) > Objects(2) Then
            P = 1
        Else
            P = 2
        End If
        If Objects(P) < Objects(3) Then
            P = 3
        End If
        If Objects(P) < Objects(4) Then
            P = 4
        End If
        Subtract = Objects(P) 'Calculate how many objects needed to be removed from the pile with the largerst number of objects.
        Objects(P) = 0
        MsgBox "I choose to remove " & Subtract & " Objects from pile " & P, , "My Choice"
    
    Else
        If Round((Pile(1, 1) + Pile(2, 1) + Pile(3, 1) + Pile(4, 1)) / 2) <> (Pile(1, 1) + Pile(2, 1) + Pile(3, 1) + Pile(4, 1)) / 2 Then
            Subtract = 8
        If Round((Pile(1, 2) + Pile(2, 2) + Pile(3, 2) + Pile(4, 2)) / 2) <> (Pile(1, 2) + Pile(2, 2) + Pile(3, 2) + Pile(4, 2)) / 2 Then
            Subtract = Subtract - 4
        End If
        If Round((Pile(1, 3) + Pile(2, 3) + Pile(3, 3) + Pile(4, 3)) / 2) <> (Pile(1, 3) + Pile(2, 3) + Pile(3, 3) + Pile(4, 3)) / 2 Then
            Subtract = Subtract - 2
        End If
        If Round((Pile(1, 4) + Pile(2, 4) + Pile(3, 4) + Pile(4, 4)) / 2) <> (Pile(1, 4) + Pile(2, 4) + Pile(3, 4) + Pile(4, 4)) / 2 Then
            Subtract = Subtract - 1
        End If
        
        If Objects(1) > Objects(2) Then
            P = 1
        Else
            P = 2
        End If
        If Objects(P) < Objects(3) Then
            P = 3
        End If
        If Objects(P) < Objects(4) Then
            P = 4
        End If
        Objects(P) = Objects(P) - Subtract
        MsgBox "I choose to remove " & Subtract & " Objects from pile " & P, , "My Choice"
    Else
        If Round((Pile(1, 2) + Pile(2, 2) + Pile(3, 2) + Pile(4, 2)) / 2) <> (Pile(1, 2) + Pile(2, 2) + Pile(3, 2) + Pile(4, 2)) / 2 Then
            Subtract = 4
            If Round((Pile(1, 3) + Pile(2, 3) + Pile(3, 3) + Pile(4, 3)) / 2) <> (Pile(1, 3) + Pile(2, 3) + Pile(3, 3) + Pile(4, 3)) / 2 Then
                Subtract = Subtract - 2
            End If
            If Round((Pile(1, 4) + Pile(2, 4) + Pile(3, 4) + Pile(4, 4)) / 2) <> (Pile(1, 4) + Pile(2, 4) + Pile(3, 4) + Pile(4, 4)) / 2 Then
                Subtract = Subtract - 1
            End If
        
            If Objects(1) > Objects(2) Then
            P = 1
            Else
            P = 2
            End If
            If Objects(P) < Objects(3) Then
                P = 3
            End If
            If Objects(P) < Objects(4) Then
                P = 4
            End If
            Objects(P) = Objects(P) - Subtract
            MsgBox "I choose to remove " & Subtract & " Objects from pile " & P, , "My Choice"
            
        Else
            If Round((Pile(1, 3) + Pile(2, 3) + Pile(3, 3) + Pile(4, 3)) / 2) <> (Pile(1, 3) + Pile(2, 3) + Pile(3, 3) + Pile(4, 3)) / 2 Then
               Subtract = 2
               If Round((Pile(1, 4) + Pile(2, 4) + Pile(3, 4) + Pile(4, 4)) / 2) <> (Pile(1, 4) + Pile(2, 4) + Pile(3, 4) + Pile(4, 4)) / 2 Then
                    Subtract = Subtract - 1
                    End If
        
                If Objects(1) > Objects(2) Then
                    P = 1
                Else
                    P = 2
                End If
                If Objects(P) < Objects(3) Then
                    P = 3
                End If
                If Objects(P) < Objects(4) Then
                    P = 4
                End If
                Objects(P) = Objects(P) - Subtract
                MsgBox "I choose to remove " & Subtract & " Objects from pile " & P, , "My Choice"
            Else
                If Objects(1) > Objects(2) Then
                    P = 1
                Else
                    P = 2
                End If
                If Objects(P) < Objects(3) Then
                    P = 3
                End If
                If Objects(P) < Objects(4) Then
                    P = 4
                End If
                Objects(P) = Objects(P) - 1
                MsgBox "I choose to remove 1 Objects from pile " & P, , "My Choice"
            End If
        End If
    End If
    End If
End If
For Q = 1 To 4
    picResult.Print "The number of objects in pile "; Q; "is "; Objects(Q)
Next Q
If Objects(1) = 0 And Objects(2) = 0 And Objects(3) = 0 And Objects(4) = 0 Then
    MsgBox "Wow, it seems that I win the game! Try again!", , "I win!"
Else
MsgBox "It's Your Turn Now!", , "Your Turn"
SubFrom = InputBox("Which pile do you want to remove objects from?", "Choose the pile")
NumSub = InputBox("How many objects do you want to remove?", "How many?")
picResult.Cls
Objects(SubFrom) = Objects(SubFrom) - NumSub
For Q = 1 To 4
    picResult.Print "The number of objects in pile "; Q; "is "; Objects(Q)
Next Q
End If
End If
End Sub

Private Sub cmdPlay_Click()
'ask the user to enter the number of objects in each pile
'this allows an experienced user to start with a position that has a winning strategy
Dim N As Integer
Dim Temp As Integer
MsgBox "There will be four piles, but you can let some of the piles be empty (set the number of objects equal zero)", , "NIM!"
MsgBox "For the game to run properly, please don't input any number that is out or range. For example, please don't try to remove any objects from an empty pile,", , "Attention"
MsgBox "Now please enter the number of objects (0 to 10) in each pile, if the number is out of range it will be automatically set to be 5.", , "NIM!"
For N = 1 To 4
        Temp = InputBox("Enter the number of objects in pile " & N, "Number of Objects in pile " & N)
    If Temp < 0 Or Temp > 10 Then
    Objects(N) = 5
    Else
    Objects(N) = Temp
    End If
    picResult.Print "The number of objects in pile "; N; "is "; Objects(N)
Next N
cmdPlay.Visible = False
cmdContinue.Visible = True
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdResource_Click()
ShellExecute Me.hWnd, "open", "http://en.wikipedia.org/wiki/Nim", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdStrategy_Click()
frmFifth.Visible = True
frmThird.Visible = False
End Sub
