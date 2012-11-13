VERSION 5.00
Begin VB.Form frmRoom5 
   BackColor       =   &H80000017&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdFour 
      BackColor       =   &H80000015&
      Caption         =   "Swap"
      Height          =   495
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdThree 
      BackColor       =   &H80000015&
      Caption         =   "Swap"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdTwo 
      BackColor       =   &H80000015&
      Caption         =   "Swap"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOne 
      BackColor       =   &H80000015&
      Caption         =   "Swap"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
   End
   Begin VB.PictureBox picFive 
      Height          =   3015
      Left            =   11640
      ScaleHeight     =   2955
      ScaleMode       =   0  'User
      ScaleWidth      =   1932.24
      TabIndex        =   8
      Top             =   1320
      Width           =   2550
   End
   Begin VB.PictureBox picFour 
      Height          =   3015
      Left            =   8880
      ScaleHeight     =   2955
      ScaleMode       =   0  'User
      ScaleWidth      =   1932.24
      TabIndex        =   7
      Top             =   1320
      Width           =   2550
   End
   Begin VB.PictureBox picThree 
      Height          =   3015
      Left            =   6120
      ScaleHeight     =   2955
      ScaleMode       =   0  'User
      ScaleWidth      =   1932.24
      TabIndex        =   6
      Top             =   1320
      Width           =   2550
   End
   Begin VB.PictureBox picTwo 
      Height          =   3015
      Left            =   3360
      ScaleHeight     =   2955
      ScaleMode       =   0  'User
      ScaleWidth      =   1932.24
      TabIndex        =   5
      Top             =   1320
      Width           =   2550
   End
   Begin VB.CommandButton cmdPuzzle 
      BackColor       =   &H80000015&
      Caption         =   "Push Button"
      Height          =   800
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   2500
   End
   Begin VB.CommandButton cmdLadder 
      BackColor       =   &H80000015&
      Caption         =   "Go Down Ladder"
      Height          =   800
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000015&
      Caption         =   "Go Back"
      Height          =   800
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   2500
   End
   Begin VB.PictureBox picOne 
      Height          =   3015
      Left            =   600
      ScaleHeight     =   2955
      ScaleMode       =   0  'User
      ScaleWidth      =   1932.24
      TabIndex        =   0
      Top             =   1320
      Width           =   2550
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000017&
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   12120
      TabIndex        =   14
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Movement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   360
      TabIndex        =   13
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label lblStoryRoom5 
      BackColor       =   &H80000017&
      Caption         =   $"frmRoom5.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   2655
      Left            =   3600
      TabIndex        =   1
      Top             =   6120
      Width           =   8175
   End
End
Attribute VB_Name = "frmRoom5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom5
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is a 'room' of the game.  It is where the user plays a sliding image game.

Option Explicit
Dim PictureNumber As Integer
Dim CTR As Integer
Dim PicName(1 To 5) As String
Dim A As Integer, B As Integer, C As Integer, D As Integer, E As Integer
Dim Swap As Integer

Private Sub cmdBack_Click()

    'User enters room 4
    frmRoom5.Visible = False
    frmRoom4.Visible = True
    
End Sub


Private Sub cmdFour_Click()
    
    'Swaps images
    Swap = D
    D = E
    E = Swap
    
    picFour.Picture = LoadPicture(App.Path & "\" & PicName(D))
    picFive.Picture = LoadPicture(App.Path & "\" & PicName(E))
    
End Sub

Private Sub cmdLadder_Click()

    'User enters room 6
    frmRoom5.Visible = False
    frmRoom6.Visible = True
    
End Sub

Private Sub cmdLookAt_Click()

    'Loads pictures into array and loads them into the picture boxes
    'Makes the four swap buttons visible
    cmdOne.Visible = True
    cmdTwo.Visible = True
    cmdThree.Visible = True
    cmdFour.Visible = True
    
    Open App.Path & "\PictureNumber.txt" For Input As #1


    CTR = 0

    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, PicName(CTR)
    Loop
    Close #1
    
        A = 3
        picOne.Picture = LoadPicture(App.Path & "\" & PicName(A))
        B = 2
        picTwo.Picture = LoadPicture(App.Path & "\" & PicName(B))
        C = 4
        picThree.Picture = LoadPicture(App.Path & "\" & PicName(C))
        D = 5
        picFour.Picture = LoadPicture(App.Path & "\" & PicName(D))
        E = 1
        picFive.Picture = LoadPicture(App.Path & "\" & PicName(E))
    
End Sub

Private Sub cmdOne_Click()

    'Swaps pictures between picOne and picTwo
    Swap = A
    A = B
    B = Swap
    
    picOne.Picture = LoadPicture(App.Path & "\" & PicName(A))
    picTwo.Picture = LoadPicture(App.Path & "\" & PicName(B))
            
End Sub

Private Sub cmdPuzzle_Click()

    'When the user thinks he/she has it right, this checks and if it is correct
    'it opens up access to room 6
    'Otherwise, it tells him/her its wrong
    If A = 1 And B = 2 And C = 3 And D = 4 And E = 5 Then
        If LadderPuzzle = False Then
            MsgBox "A ladder appears in the floor.  Also, a bunch of coins fell out from behind the tablets.", , ""
            Coins = Coins + 20
            LadderPuzzle = True
            cmdLadder.Visible = True
            cmdPuzzle.Visible = False
            
        End If
    Else
        MsgBox "That's not right.", , ""
        
    End If
    
End Sub

Private Sub cmdThree_Click()
    
    'Swaps pictures in picThree and picFour
    Swap = C
    C = D
    D = Swap
    
    picThree.Picture = LoadPicture(App.Path & "\" & PicName(C))
    picFour.Picture = LoadPicture(App.Path & "\" & PicName(D))
    
End Sub

Private Sub cmdTwo_Click()
    
    'Swaps pictures in picTwo and picThree
    Swap = B
    B = C
    C = Swap
    
    picTwo.Picture = LoadPicture(App.Path & "\" & PicName(B))
    picThree.Picture = LoadPicture(App.Path & "\" & PicName(C))
    
End Sub

Private Sub Form_Load()

    'Loads pictures into array and loads them into the picture boxes
    'Makes the four swap buttons visible
    
    Open App.Path & "\PictureNumber.txt" For Input As #1


    CTR = 0

    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, PicName(CTR)
    Loop
    Close #1
    
        A = 3
        picOne.Picture = LoadPicture(App.Path & "\" & PicName(A))
        B = 2
        picTwo.Picture = LoadPicture(App.Path & "\" & PicName(B))
        C = 4
        picThree.Picture = LoadPicture(App.Path & "\" & PicName(C))
        D = 5
        picFour.Picture = LoadPicture(App.Path & "\" & PicName(D))
        E = 1
        picFive.Picture = LoadPicture(App.Path & "\" & PicName(E))
        
End Sub
