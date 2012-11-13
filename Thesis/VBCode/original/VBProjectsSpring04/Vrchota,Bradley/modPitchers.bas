Attribute VB_Name = "modPitchers"
'Project Name: MLBPitchers (MLBPitchers.vbp)
'Form Name: modPitchers (modPitchers.bas)
'Author: Bradley Vrchota
'Date: March 14, 2004
'Purpose: The purpose of the module is to dimension variables and
        'arrays and make them available to all forms

'Make the 5 arrays and the counter available to all forms
Public pitcher(1 To 20) As String
Public wins(1 To 20) As Integer
Public losses(1 To 20) As Integer
Public ERA(1 To 20) As Single
Public strikeouts(1 To 20) As Integer
Public ctr As Integer
Public J As Integer

'dim path of pitcher stat file and make available to all forms
Public PATH As String
    
    

