VERSION 5.00
Begin VB.Form frmSecondForm 
   BackColor       =   &H00FF0000&
   Caption         =   "Rugby Stats"
   ClientHeight    =   7200
   ClientLeft      =   2310
   ClientTop       =   2130
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10575
   Begin VB.CommandButton cmdWeight 
      Caption         =   "Sort Weight"
      Height          =   855
      Left            =   5640
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdTall 
      Caption         =   "Sort Height"
      Height          =   855
      Left            =   4080
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   855
      Left            =   2400
      TabIndex        =   4
      Top             =   5760
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   4575
      Left            =   720
      ScaleHeight     =   4515
      ScaleWidth      =   7515
      TabIndex        =   3
      Top             =   360
      Width           =   7575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start Search"
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   735
      Left            =   7320
      TabIndex        =   1
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   8880
      TabIndex        =   0
      Top             =   5880
      Width           =   1335
   End
End
Attribute VB_Name = "frmSecondForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Counter As Integer          'Declaring variables
Dim Pos As Integer
Dim A As String
Dim G As Integer
Dim Pass As Integer
Dim Size As Integer
Dim T As Integer
Dim Z As Integer
Dim Names(1 To 15) As String
Dim Tall(1 To 15) As Single
Dim Weight(1 To 15) As Single
Dim TempNames As String
Dim TempTall As Single
Dim TempWeight As Integer
Dim Found As Boolean


Private Sub cmdBack_Click()
    frmSecondForm.Hide      'removes form from screen
    frmFirstForm.Show       'pulls next form up
End Sub

Private Sub cmdQuit_Click()
    frmSecondForm.Hide      'removes form from screen
    frmQuit.Show            'pulls next form up
End Sub
Private Sub cmdSearch_Click()
    Counter = 0                     'setting variables before their use
    G = 0
    Found = False
    Open App.Path & "\RugbyStats.txt" For Input As #1   'opens up file
    Do Until EOF(1)
        Counter = Counter + 1
        Input #1, Names(Counter), Tall(Counter), Weight(Counter)
    Loop
    Close #1                                            'closes file
    A = InputBox("Enter a name", "Rugby players")       'sets input as "A"
    Do While Found = False
        G = G + 1
        If Names(G) = A Then
            Found = True
        ElseIf Names(G) = Not A Then
            Found = False
        End If
    Loop
    If Found = True Then
        picResults.Print "Name "; Tab(25); "Height in inches "; Tab(45); "Weight "   'displays text and spacing onto picture box
        picResults.Print A, Tab(25); Tall(Counter), Tab(45); Weight(Counter)
    Else
        MsgBox "He does not play rugby for SJU.", "Error!"      'opens message box
    End If

End Sub

Private Sub cmdLoad_Click()
    picResults.Cls                                      'clears picture box
    Size = 0
    Open App.Path & "\RugbyStats.txt." For Input As #1
    Do Until EOF(1)
        Size = Size + 1
        Input #1, Names(Size), Tall(Size), Weight(Size)
    Loop
    Close #1
    picResults.Print "Name", Tab(25); "Height in inches"; Tab(45); "Weight"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab(25); Tall(Pos); Tab(45); Weight(Pos)
    Next Pos
    
    
  
End Sub

Private Sub cmdTall_Click()
                                        'this is sorting by height.
    picResults.Cls
    T = 0
    For Pass = 1 To Size - 1
        For T = 1 To Size - Pass
            If Tall(T + 1) > Tall(T) Then
                TempTall = Tall(T)
                Tall(T + 1) = TempTall
                Names(T + 1) = Names(T)
                TempNames = Names(T)
                Names(T) = Names(T + 1)
                Names(T + 1) = TempNames
                TempWeight = Weight(T)
                Weight(T) = Weight(T + 1)
                Weight(T + 1) = TempWeight
            End If
        Next T
    Next Pass
    picResults.Print "Name", Tab(25); "Height in inches"; Tab(45); "Weight"
    For Z = 1 To Size
        picResults.Print Names(Z); Tab(25); Tall(Z); Tab(45); Weight(Z)
    Next Z
End Sub

Private Sub cmdWeight_Click()
                                        'this is the sorting by weight.
    picResults.Cls
    T = 0
    For Pass = 1 To Size - 1
        For T = 1 To Size - Pass
            If Weight(T + 1) > Weight(T) Then
                TempWeight = Weight(T)
                Weight(T + 1) = TempWeight
                Names(T + 1) = Names(T)
                TempNames = Names(T)
                Names(T) = Names(T + 1)
                Names(T + 1) = TempNames
                Tall(T + 1) = Tall(T)
                TempTall = Tall(T)
                Tall(T + 1) = TempTall
            End If
        Next T
    Next Pass
    picResults.Print "Name", Tab(25); "Height in inches"; Tab(45); "Weight"
    For Z = 1 To Size
        picResults.Print Names(Z); Tab(25); Tall(Z); Tab(45); Weight(Z)
    Next Z
End Sub

                                                                        'Eric Bendzinski Project 1.vbp
                                                                        'frmSecondForm
                                                                        'Eric Bendzinski
                                                                        'Written 11/1/06 and 11/3/06
                                                                        
