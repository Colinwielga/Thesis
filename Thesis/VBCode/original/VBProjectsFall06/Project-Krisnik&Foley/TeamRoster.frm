VERSION 5.00
Begin VB.Form TeamRoster 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Team Roster"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Team Roster"
   ScaleHeight     =   7785
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   240
      Picture         =   "TeamRoster.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortByWeight 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort Players By Weight"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortByHeight 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort Players By Height"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdSortByNumber 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort Players By Number"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080C0FF&
      Caption         =   "Search For Player"
      Height          =   735
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Return To Homepage"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdShowRoster 
      BackColor       =   &H0080C0FF&
      Caption         =   "Show Twins Roster"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   7335
      Left            =   3000
      ScaleHeight     =   7275
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "TeamRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name: Twins Baseball
' Form Name: Homepage
' Authors: Jake Krisnik & Mike Foley
' Date Written: October 21, 2006
' Form Objective: To provide the user with information regarding the Twins active roster.
'                 We intend to give players numbers, names, weights, heights, and positions.
'                 The user will also be able to sort this information according to different
'                 criterion, as well as be able to search to see if a particular player is
'                 a member of the Minnesota Twins.

Option Explicit
    Dim PlayerNumber(1 To 100) As Integer
    Dim PlayerWeight(1 To 100) As Integer
    Dim PlayerHeight(1 To 100) As Integer
    Dim PlayerName(1 To 100) As String
    Dim PlayerPosition(1 To 100) As String
    Dim Ctr As Integer
' This code allows the user to return to the Homepage form and hides the Team Roster form.
Private Sub cmdReturn_Click()
    HomePage.Show
    TeamRoster.Hide
End Sub

Private Sub cmdSearch_Click()
' This code searches through our array to see if the name the user enters in the input box
' matches with the name of a Twins player on the active roster. If it does, the message
' prompt will tell the user that that individual is on the team and vise versa.
    Dim Player As String
    Dim I As Single
    Dim Found As Boolean    ' using the boolean variable Found we are able to match the users input with our existing file
    Player = InputBox("Enter Players First And Last Name (**No Comma**)", Player)
    I = 0
    Found = False
    Do While I < 26
        I = I + 1
        If Player = PlayerName(I) Then
            MsgBox Player & " " & "is a member of the Minnesota Twins."
            Found = True
        End If
    Loop
    If Found = False Then
        MsgBox Player & " " & "is not a member of the Minnesota Twins"
    End If

End Sub

Private Sub cmdShowRoster_Click()
' The code for this command button is used to access our text file which includes player
' information. It prints out the file in a user friendly format which they can then look
' through to gain information on particular players of the Minnesota Twins.

      Open App.Path & "\Team.txt" For Input As #1
    Ctr = 0
        picResults.Cls
        picResults.Print "                                          Minnesota Twins Active Roster                   "
        picResults.Print "___________________________________________________________________________________"
        picResults.Print "Name                                      Weight             Height            Number           Position"
        picResults.Print "                                               (Lbs.)                (In.)  "
        picResults.Print "___________________________________________________________________________________"
    
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, PlayerNumber(Ctr), PlayerName(Ctr), PlayerHeight(Ctr), PlayerWeight(Ctr), PlayerPosition(Ctr)
        picResults.Print PlayerName(Ctr); "          ", PlayerWeight(Ctr), PlayerHeight(Ctr), PlayerNumber(Ctr), UCase(PlayerPosition(Ctr))
    Loop
        picResults.Print "************************************************************************************************************************"
    Close #1
End Sub
    

Private Sub cmdSortByHeight_Click()
' This code uses a temporary variable to sort the array according to height in ascending order.
' This code is explained in further detail on the Project Writeup.
    Dim Pass As Integer
    Dim Temp1 As Integer
    Dim Temp2 As String
    Dim Temp3 As Integer
    Dim Temp4 As Integer
    Dim Temp5 As String
    Dim I As Integer
        picResults.Cls
        picResults.Print "                                          Minnesota Twins Active Roster                   "
        picResults.Print "_________________________________________________________________________"
        picResults.Print "Height            Name                                        Number         Weight              Position"
        picResults.Print "(In.)                                                                                        (Lbs.)"
        picResults.Print "_________________________________________________________________________"
    For Pass = 1 To (Ctr - 1)
        For I = 1 To Ctr - Pass
            If PlayerHeight(I) > PlayerHeight(I + 1) Then
                Temp1 = PlayerHeight(I)
                PlayerHeight(I) = PlayerHeight(I + 1)
                PlayerHeight(I + 1) = Temp1
                Temp2 = PlayerName(I)
                PlayerName(I) = PlayerName(I + 1)
                PlayerName(I + 1) = Temp2
                Temp3 = PlayerNumber(I)
                PlayerNumber(I) = PlayerNumber(I + 1)
                PlayerNumber(I + 1) = Temp3
                Temp4 = PlayerWeight(I)
                PlayerWeight(I) = PlayerWeight(I + 1)
                PlayerWeight(I + 1) = Temp4
                Temp5 = PlayerPosition(I)
                PlayerPosition(I) = PlayerPosition(I + 1)
                PlayerPosition(I + 1) = Temp5
            End If
        Next I
    Next Pass
        For I = 1 To Ctr
            picResults.Print PlayerHeight(I), PlayerName(I), Tab(45); PlayerNumber(I), PlayerWeight(I), UCase(PlayerPosition(I))
        Next I
            picResults.Print "**************************************************************************************************************"
End Sub


Private Sub cmdSortByNumber_Click()
' This code uses a temporary variable to sort the array by player number in ascending order.
    Dim Pass As Integer
    Dim Temp1 As Integer
    Dim Temp2 As String
    Dim Temp3 As Integer
    Dim Temp4 As Integer
    Dim Temp5 As String
    Dim I As Integer
        picResults.Cls
        picResults.Print "                                          Minnesota Twins Active Roster                   "
        picResults.Print "_________________________________________________________________________"
        picResults.Print "Number            Name                                        Height         Weight              Position"
        picResults.Print "                                                                          (In.)             (Lbs.)"
        picResults.Print "_________________________________________________________________________"
    For Pass = 1 To (Ctr - 1)
        For I = 1 To Ctr - Pass
            If PlayerNumber(I) > PlayerNumber(I + 1) Then
                Temp1 = PlayerNumber(I)
                PlayerNumber(I) = PlayerNumber(I + 1)
                PlayerNumber(I + 1) = Temp1
                Temp2 = PlayerName(I)
                PlayerName(I) = PlayerName(I + 1)
                PlayerName(I + 1) = Temp2
                Temp3 = PlayerHeight(I)
                PlayerHeight(I) = PlayerHeight(I + 1)
                PlayerHeight(I + 1) = Temp3
                Temp4 = PlayerWeight(I)
                PlayerWeight(I) = PlayerWeight(I + 1)
                PlayerWeight(I + 1) = Temp4
                Temp5 = PlayerPosition(I)
                PlayerPosition(I) = PlayerPosition(I + 1)
                PlayerPosition(I + 1) = Temp5
            End If
        Next I
    Next Pass
        For I = 1 To Ctr
            picResults.Print PlayerNumber(I), PlayerName(I), Tab(45); PlayerHeight(I), PlayerWeight(I), UCase(PlayerPosition(I))
        Next I
            picResults.Print "**************************************************************************************************************"
End Sub

Private Sub cmdSortByWeight_Click()
' This code uses a temporary variable to sort the array by weight in ascending order.
Dim Pass As Integer
    Dim Temp1 As Integer
    Dim Temp2 As String
    Dim Temp3 As Integer
    Dim Temp4 As Integer
    Dim Temp5 As String
    Dim I As Integer
        picResults.Cls
        picResults.Print "                                          Minnesota Twins Active Roster                   "
        picResults.Print "_________________________________________________________________________"
        picResults.Print "Weight            Name                                        Number         Height              Position"
        picResults.Print "(Lbs.)                                                                                        (In.)"
        picResults.Print "_________________________________________________________________________"
    For Pass = 1 To (Ctr - 1)
        For I = 1 To Ctr - Pass
            If PlayerWeight(I) > PlayerWeight(I + 1) Then
                Temp1 = PlayerWeight(I)
                PlayerWeight(I) = PlayerWeight(I + 1)
                PlayerWeight(I + 1) = Temp1
                Temp2 = PlayerName(I)
                PlayerName(I) = PlayerName(I + 1)
                PlayerName(I + 1) = Temp2
                Temp3 = PlayerNumber(I)
                PlayerNumber(I) = PlayerNumber(I + 1)
                PlayerNumber(I + 1) = Temp3
                Temp4 = PlayerHeight(I)
                PlayerHeight(I) = PlayerHeight(I + 1)
                PlayerHeight(I + 1) = Temp4
                Temp5 = PlayerPosition(I)
                PlayerPosition(I) = PlayerPosition(I + 1)
                PlayerPosition(I + 1) = Temp5
            End If
        Next I
    Next Pass
        For I = 1 To Ctr
            picResults.Print PlayerWeight(I), PlayerName(I), Tab(45); PlayerNumber(I), PlayerHeight(I), UCase(PlayerPosition(I))
        Next I
            picResults.Print "**************************************************************************************************************"
End Sub
