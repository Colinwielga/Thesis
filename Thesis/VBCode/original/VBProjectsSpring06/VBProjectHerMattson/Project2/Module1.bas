Attribute VB_Name = "Module1"
'AmazingQuiz
'Module1
'Ee Her and Jennifer Mattson
'Written on 3/18/06
'These variable are publized because they are used in more then one form.

'counter keeps track of how many questions the user answer correctly.
Public Counter As Integer
'Size keeps track of how large the arrays are.
Public Size As Integer
'These variables deal with personal information.
Public userlastname(1 To 100) As String
Public userfirstname(1 To 100) As String
Public usergender(1 To 100) As String
Public userage(1 To 100) As Integer
Public userscore(1 To 100) As Integer
'These two variables keep track of time.
Public elapsedTime As Double
Public run As Boolean

