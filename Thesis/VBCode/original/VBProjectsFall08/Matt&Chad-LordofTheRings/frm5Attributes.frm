VERSION 5.00
Begin VB.Form frm5Attributes 
   Caption         =   "Form3"
   ClientHeight    =   7785
   ClientLeft      =   1440
   ClientTop       =   1515
   ClientWidth     =   10395
   LinkTopic       =   "Form3"
   Picture         =   "frm5Attributes.frx":0000
   ScaleHeight     =   7785
   ScaleWidth      =   10395
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Journey"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      Height          =   255
      Left            =   9120
      TabIndex        =   5
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Information"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortAgility 
      Caption         =   "Sort By Agility"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdStrength 
      Caption         =   "Sort by Strength"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortWisdown 
      Caption         =   "Sort By Wisdom"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   3360
      ScaleHeight     =   2475
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   4920
      Width           =   3975
   End
End
Attribute VB_Name = "frm5Attributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim CTR As Integer, Names(1 To 10) As String, Agility(1 To 10) As Integer
    Dim Wisdom(1 To 10) As Integer, Strength(1 To 10) As Integer

Private Sub cmdGoBack_Click()
    frm5Attributes.Hide
    frm2Characters.Show
End Sub

Private Sub cmdLoad_Click()
    picResults.Cls
    Dim J As Integer
    Open App.Path & "\Attributes.txt" For Input As #1
        CTR = 0
            Do Until EOF(1)
                CTR = CTR + 1
                Input #1, Names(CTR), Wisdom(CTR), Agility(CTR), Strength(CTR)
            Loop
        Close #1
    picResults.Print "Names", "Wisdom", "Agility", "Strength"
    picResults.Print "****************************************************************"
    For J = 1 To CTR
        picResults.Print Names(J), Wisdom(J), Agility(J), Strength(J)
    Next J
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSortAgility_Click()
Dim Pass As Integer, Pos As Integer, n As Integer
Dim TempNames As String, TempWisdom As Integer, TempAgility As Integer, TempStrength As Integer

picResults.Cls
    For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Agility(Pos) < Agility(Pos + 1) Then
                    TempWisdom = Wisdom(Pos)
                    Wisdom(Pos) = Wisdom(Pos + 1)
                    Wisdom(Pos + 1) = TempWisdom
                    
                    TempAgility = Agility(Pos)
                    Agility(Pos) = Agility(Pos + 1)
                    Agility(Pos + 1) = TempAgility
                    
                    TempNames = Names(Pos)
                    Names(Pos) = Names(Pos + 1)
                    Names(Pos + 1) = TempNames
                    
                    TempStrength = Strength(Pos)
                    Strength(Pos) = Strength(Pos + 1)
                    Strength(Pos + 1) = TempStrength
                End If
            Next Pos
        Next Pass
        
    picResults.Print "Names", "Wisdom", "Agility", "Strength"
    picResults.Print "****************************************************************"
    For n = 1 To CTR
        picResults.Print Names(n), Wisdom(n), Agility(n), Strength(n)
    Next n
End Sub

Private Sub cmdSortWisdown_Click()
Dim Pass As Integer, Pos As Integer, n As Integer
Dim TempNames As String, TempWisdom As Integer, TempAgility As Integer, TempStrength As Integer

picResults.Cls
    For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Wisdom(Pos) < Wisdom(Pos + 1) Then
                    TempWisdom = Wisdom(Pos)
                    Wisdom(Pos) = Wisdom(Pos + 1)
                    Wisdom(Pos + 1) = TempWisdom
                    
                    TempAgility = Agility(Pos)
                    Agility(Pos) = Agility(Pos + 1)
                    Agility(Pos + 1) = TempAgility
                    
                    TempNames = Names(Pos)
                    Names(Pos) = Names(Pos + 1)
                    Names(Pos + 1) = TempNames
                    
                    TempStrength = Strength(Pos)
                    Strength(Pos) = Strength(Pos + 1)
                    Strength(Pos + 1) = TempStrength
                End If
            Next Pos
        Next Pass
        
    picResults.Print "Names", "Wisdom", "Agility", "Strength"
    picResults.Print "****************************************************************"
    For n = 1 To CTR
        picResults.Print Names(n), Wisdom(n), Agility(n), Strength(n)
    Next n
    
End Sub
Private Sub cmdStrength_Click()
Dim Pass As Integer, Pos As Integer, n As Integer
Dim TempNames As String, TempWisdom As Integer, TempAgility As Integer, TempStrength As Integer

picResults.Cls
    For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Strength(Pos) < Strength(Pos + 1) Then
                    TempWisdom = Wisdom(Pos)
                    Wisdom(Pos) = Wisdom(Pos + 1)
                    Wisdom(Pos + 1) = TempWisdom
                    
                    TempAgility = Agility(Pos)
                    Agility(Pos) = Agility(Pos + 1)
                    Agility(Pos + 1) = TempAgility
                    
                    TempNames = Names(Pos)
                    Names(Pos) = Names(Pos + 1)
                    Names(Pos + 1) = TempNames
                    
                    TempStrength = Strength(Pos)
                    Strength(Pos) = Strength(Pos + 1)
                    Strength(Pos + 1) = TempStrength
                End If
            Next Pos
        Next Pass
        
    picResults.Print "Names", "Wisdom", "Agility", "Strength"
    picResults.Print "****************************************************************"
    For n = 1 To CTR
        picResults.Print Names(n), Wisdom(n), Agility(n), Strength(n)
    Next n
End Sub
