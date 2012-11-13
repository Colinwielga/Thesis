Attribute VB_Name = "LBL_ProBankMod"
Option Explicit
'Levi Lowell
'CSCI 130
'Professor: Imad rahal
'March 25th, 2007
'This form allows certain variables to be globally set and used in multiple forms, particularly
'in regards to my Login form and the entire login process within the piece.  Much of the code
'pertaining to my Login access application was borrowed from Bill Macy and his exceptional program
'"Mario Madness".  Other code was borrowed from my previous work in my labs and exercises.  This
'information will be cited in my citations application within the Form frmMeetme.  Enjoy.

Global Number As Integer        'Sets global variables usable by multiple forms within the project
Global username As String
Global Password As String
Global inputName As String
Global inputPassword As String
Global wrongEntry As Boolean
Global chosenname As String
Global chosenpassword As String
Global Pos As Boolean



