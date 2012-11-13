Attribute VB_Name = "statisticsmodule"
'Project Name : Basketball Game Statistics (BasketballPlayersInput.vbp)
'Form Name : statisticsmodule (Module1.bas)
'Author : Anthony Mayerhofer
'Date Written : March 15, 2004
'Purpose of Project : To read the file for the statistics
                      'from the basketball game
                      'then have the user predict
                      'the most efficient shooter
                      'and finally display the results
                      'of several comparisons
                      'to evaluate players on several criteria

' Purpose of code module: dimension all variables and arrays
                        'that will need to be used on multiple forms

Public CTR As Integer, X As Integer
Public names(1 To 5) As String
Public points(1 To 5) As Integer
Public shootingpercentage(1 To 5) As Single
Public rebounds(1 To 5) As Integer

