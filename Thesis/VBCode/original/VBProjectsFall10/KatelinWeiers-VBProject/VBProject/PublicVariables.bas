Attribute VB_Name = "modVariables"
Option Explicit 'must declare all variables

'declare public variables relating to the individual players
Public PlayerNumber(1 To 50) As Long, FirstName(1 To 50) As String, LastName(1 To 50) As String
Public Position(1 To 50) As String, PlayerBattingAvg(1 To 50) As Single, Birthdate(1 To 50) As String
Public Ctr As Integer

'declare public variables relating to the overall team record
Public Season(1 To 100) As Long, Wins(1 To 100) As Long, Losses(1 To 100) As Long
Public Attendance(1 To 100) As Long, Champions(1 To 100) As String
Public CtrTeam As Integer
