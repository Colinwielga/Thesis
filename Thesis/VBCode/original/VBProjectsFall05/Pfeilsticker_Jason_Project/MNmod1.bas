Attribute VB_Name = "Startup"
'Project name: MNTWINS(stats)
'Module: Startup(MNmod1)
'Author: Jason Pfeilsticker
'Date Written: October 30, 2005
'Objective:  The objective is this module is to make public the variables
'that will be used throughout the project.  The main variable in here are
'the arrays and variable used to find a player in a file.

Public Pitchers(1 To 20) As String, Batters(1 To 25) As String
Public Wins(1 To 20) As Integer, Losses(1 To 20) As Integer, Saves(1 To 20) As Integer, Strikeouts(1 To 20) As Integer
Public ERA(1 To 20) As Single, Innings(1 To 20) As String
Public Games(1 To 25) As Integer, AtBats(1 To 25) As Integer, Hits(1 To 25) As Integer, HR(1 To 25) As Integer, RBI(1 To 25) As Integer
Public AVG(1 To 25) As String
Public CTR As Integer
Public CTR2 As Integer
'declares what will initally be the first arrays

Public Name As String
Public I As Integer
Public Found As Boolean

'declare variables used to find a players
Sub main()
'This will have the project start on the Batters form
PositionPlayers.Show
End Sub
