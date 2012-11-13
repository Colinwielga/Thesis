Attribute VB_Name = "Module1"
Option Explicit
Public Numbers(1 To 26) As Single
Public Checks(1 To 26) As Single
Public Ctr As Double, J As Integer
Public CaseDollar(1 To 26) As Single
Public r As Double
Public Ctr3 As Single
Public Sum As Single
Public K As Integer
Public id As String
Public amount As Single
Public Num As Integer
Public Good As Single
Public Declare Function PlaySound Lib "winmm.dll" Alias _
"sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As _
Long) As Long
'Project: Deal or No Deal
'Module
'Holly Reinking and Danielle Karp
'Written 3/15/09
'Purpose: To load and randomly assort an array for the project, to dim variables publically for the project and to change forms

 
'http://soundboard.com/sb/Deal_Or_No_Deal_sounds.aspx is where we found all of our sound clips

 Sub Main()
    
Ctr = 0
Open App.Path & "\dealornodeal.txt" For Input As #1 'This code enters in all of the numbers in our file as an array
    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Numbers(Ctr)
    Loop
Close #1
    

Randomize
Ctr = 0
Do While Ctr < 26
r = Int(26 * Rnd) + 1
    If Checks(r) = 0 Then    'This code randomizes our array numbers so that each case number will have a different amount assigned to it everytime the game is played
        Checks(r) = r
        Ctr = Ctr + 1
    CaseDollar(Ctr) = Numbers(r)
    End If
Loop

frm1welcome.Show             'Shows one form and Hides another
frmdealornodeal.Hide

End Sub
