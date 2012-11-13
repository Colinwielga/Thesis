VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H80000003&
   Caption         =   "Search for Player"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFindBatter 
      Caption         =   "Search for Position Player"
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdFindPitcher 
      Caption         =   "Search for Pitcher"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturnSearch 
      Caption         =   "Return to Main Page"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   1575
      Left            =   360
      ScaleHeight     =   1515
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   960
      Width           =   6735
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form allows the user to search for a particular player

Private Sub cmdFindBatter_Click()
    'Batter subroutine
    'Declaring variables
    Dim userbat As String, found As Boolean, pos As Integer
    picResults.Cls
    
    found = False
    pos = 0
    'Receive name from input box and search for player
    userbat = InputBox("Enter the name of the batter you're searching for", "Search for batter")
    Do While (found = False And pos < bcounter)
        pos = pos + 1
        If Batters(pos) = userbat Then
            found = True
        End If
    Loop
    'Determines worthiness and prints message in picture box
    If found = True Then
        If BatterPos(pos) = 2 Then
            If BatterTotals(pos) > 2200 Then
                picResults.Print userbat & " is a Hall of Fame catcher with a score of " & BatterTotals(pos)
            Else
                picResults.Print userbat & " is not a Hall fo Fame catcher because his score of "; BatterTotals(pos) & " is too small."
            End If
        ElseIf BatterPos(pos) = 3 Then
            If BatterTotals(pos) > 2500 Then
                picResults.Print userbat & " is a Hall of Fame first baseman with a score of " & BatterTotals(pos)
            Else
                picResults.Print userbat & " is not a Hall fo Fame first baseman because his score of "; BatterTotals(pos) & " is too small."
            End If
        ElseIf BatterPos(pos) = 4 Then
            If BatterTotals(pos) > 2300 Then
                picResults.Print userbat & " is a Hall of Fame second baseman with a score of " & BatterTotals(pos)
            Else
                picResults.Print userbat & " is not a Hall fo Fame second baseman because his score of "; BatterTotals(pos) & " is too small."
            End If
        ElseIf BatterPos(pos) = 5 Then
            If BatterTotals(pos) > 2350 Then
                picResults.Print userbat & " is a Hall of Fame third baseman with a score of " & BatterTotals(pos)
            Else
                picResults.Print userbat & " is not a Hall fo Fame third baseman because his score of "; BatterTotals(pos) & " is too small."
            End If
        ElseIf BatterPos(pos) = 6 Then
            If BatterTotals(pos) > 2200 Then
                picResults.Print userbat & " is a Hall of Fame shortstop with a score of " & BatterTotals(pos)
            Else
                picResults.Print userbat & " is not a Hall fo Fame shortstop because his score of "; BatterTotals(pos) & " is too small."
            End If
        ElseIf BatterPos(pos) = 7 Then
            If BatterTotals(pos) > 2450 Then
                picResults.Print userbat & " is a Hall of Fame corner outfielder with a score of " & BatterTotals(pos)
            Else
                picResults.Print userbat & " is not a Hall fo Fame corner outfielder because his score of "; BatterTotals(pos) & " is too small."
            End If
        ElseIf BatterPos(pos) = 8 Then
            If BatterTotals(pos) > 2300 Then
                picResults.Print userbat & " is a Hall of Fame center fielder with a score of " & BatterTotals(pos)
            Else
                picResults.Print userbat & " is not a Hall fo Fame center fielder because his score of "; BatterTotals(pos) & " is too small."
            End If
        End If
    ElseIf found = False Then
        MsgBox userbat & " could not be found", , "Error"
    End If
End Sub

Private Sub cmdFindPitcher_Click()
    'Pitcher subroutine
    'Declaring variables
    Dim Userpitch As String, found As Boolean, pos As Integer
    picResults.Cls
    
    found = False
    pos = 0
    'Receives name from input box and searches for it
    Userpitch = InputBox("Enter the name of the pitcher you're searching for", "Search for Pitcher")
    Do While (found = False And pos < pcounter)
        pos = pos + 1
        If Pitchers(pos) = Userpitch Then
            found = True
        End If
    Loop
    'Determines worthiness and prints message in picture box
    If found = True Then
        If PitcherPos(pos) = 1 Then
            If PitcherTotals(pos) >= 2300 Then
                picResults.Print Userpitch & " is a Hall of Famer with a score of " & PitcherTotals(pos)
            Else
                picResults.Print Userpitch & " is not a Hall of Famer, because his score of "; PitcherTotals(pos) & " is too small."
            End If
        ElseIf PitcherPos(pos) = 2 Then
            If PitcherTotals(pos) >= 1450 Then
                picResults.Print Userpitch & " is a Hall of Famer with a score of " & PitcherTotals(pos)
            Else
                picResults.Print Userpitch & " is not a Hall of Famer, because his score of "; PitcherTotals(pos) & " is too small."
            End If
        End If
    ElseIf found = False Then
        MsgBox Userpitch & " could not be found", , "Error"
    End If
End Sub

Private Sub cmdReturnSearch_Click()
    'Makes home page appear
    frmSearch.Visible = False
    frmExplanation.Visible = False
    frmRankings.Visible = False
    frmHome.Visible = True
    
End Sub
