Attribute VB_Name = "Module1"
Public CardPix(4 To 56) As String, dublinpix(1 To 5) As String, gettysburgpix(1 To 3) As String, airbornepix(1 To 2) As String, dcltpix(1 To 2) As String, hargravepix(1 To 3) As String, ldacpix(1 To 3) As String, vegaspix(1) As String 'declares all picture array text files
Public Cash As Integer  'Declares a variable for the black jack form

Sub main()
Cash = 100 'Sets cash to 100 every time a user opens this file for the black jack form

'Open some files of picture names and put them in several arrays
Dim CTR As Integer
Open App.Path & "\cardnames.txt" For Input As #1       'opens a file
CTR = 3
Do While Not EOF(1)                                    'reads file until no more information is available
    CTR = CTR + 1
    Input #1, CardPix(CTR)
Loop
Close #1                                               'closes a file


Open App.Path & "\dublinpix.txt" For Input As #3       'opens a file
CTR = 0
Do While Not EOF(3)                                    'reads file until no more information is available
    CTR = CTR + 1
    Input #3, dublinpix(CTR)
Loop
Close #3                                               'closes a file


Open App.Path & "\gettysburgpix.txt" For Input As #4   'opens a file
CTR = 0
Do While Not EOF(4)                                    'reads file until no more information is available
    CTR = CTR + 1
    Input #4, gettysburgpix(CTR), hargravepix(CTR), ldacpix(CTR)
Loop
Close #4                                               'closes a file


Open App.Path & "\AirborneDCLT.txt" For Input As #5    'opens a file
CTR = 0
Do While Not EOF(5)                                    'reads file until no more information is available
    CTR = CTR + 1
    Input #5, airbornepix(CTR), dcltpix(CTR)
Loop
Close #5                                               'closes a file

Open App.Path & "\vegas.txt" For Input As #6           'opens a file
CTR = 1
Input #6, vegaspix(CTR)                                'loads a single picture
Close #6


'display the startup form
frmMain.Show                                           'goes to main form

End Sub

