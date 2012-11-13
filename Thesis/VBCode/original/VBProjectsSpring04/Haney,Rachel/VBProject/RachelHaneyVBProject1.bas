Attribute VB_Name = "RachelHaney8"
'PlanYourVacation (RachelHaneyVBProject8)
'RachelHaney8 (RachelHaneyVBProject1.bas)
'Rachel Haney 3/11/04
'The purpose of this module is to make
'several variables available to more than
'one form.  It will also load a piece of
'code that is to appear on a form when it
'is loaded later in the program.
Public CTR As Integer
Public Total As Single
Public Spend As Single
Public People As Integer
Public City As Integer
Public Travel As Integer
Public Visit As Integer
Public Room As Integer
Public Vehicle(1 To 5) As String
Public Price(1 To 5) As Single
Public tempVehicle As String
Public tempPrice As Single
Public PASS As Integer
Public COMP As Integer
Public J As Integer

Sub Main()
'this code allows the information to be put into an array
'as soon as the user moves to this form.  it also displays
'the information so the user knows how much each of their
'options will cost

'initialize CTR to zero, to be used for position in the array
    Dim PATH As String
    PATH = "N:\CS130\handin\Haney,Rachel\VBProject\"
    CTR = 0
'Prepare the file to be read
'Open "M:\CS130\Rachel Haney\RachelHaneyVBProjectNote.txt" For Input As #1
    Open PATH + "RachelHaneyVBProjectNote.txt" For Input As #1
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, Vehicle(CTR), Price(CTR)
        Loop

'use Bubble Sort to arrange transportation in price order.
    For PASS = 1 To CTR - 1
        For COMP = 1 To CTR - PASS
            If Price(COMP) < Price(COMP + 1) Then
        
                'switch price
                tempPrice = Price(COMP)
                Price(COMP) = Price(COMP + 1)
                Price(COMP + 1) = tempPrice
            
                'and also vehicle
                tempVehicle = Vehicle(COMP)
                Vehicle(COMP) = Vehicle(COMP + 1)
                Vehicle(COMP + 1) = tempVehicle
            
            End If
        Next COMP
    Next PASS
    Close #1
'Displays the first form for the user to start.
    RachelHaney1.Show
End Sub

