Attribute VB_Name = "mdlPublicSubs"
Option Explicit
'This module containt all of the Public Subroutines used more than once throughout the program

'Public Subroutine which reads the login data and  score data into array for use by multiple forms and buttons (scores.txt and loginData.txt ought to be concurrent with eachother)
Public Sub ReadLogin()
    'Opens files
    Open App.Path & "\data\LoginData.txt" For Input As #1
    Open App.Path & "\Data\Scores.txt" For Input As #2
        'Initiates variable
        loginCtr = 0
        'loops to read the files
        Do Until EOF(1)
            loginCtr = loginCtr + 1
            Input #1, userName(loginCtr), firstName(loginCtr), lastName(loginCtr), password(loginCtr), ClassEnrolled(loginCtr)
            Input #2, studentGradeName(loginCtr), StudentGrade(loginCtr), studentCorrect(loginCtr), studentWrong(loginCtr), StudentAttempted(loginCtr)
        Loop
    
    Close #2
    Close #1
End Sub
'Public Subroutine which reads the classlist data file and places it into arrays for use by multiple forms
Public Sub ReadClasses()
    'Opens the classlist text file
    Open App.Path & "\data\ClassList.txt" For Input As #1
        'Initiates the varaible
        classCtr = 0
        'Loops until the end of the file to read it into the array
        Do Until EOF(1)
            classCtr = classCtr + 1
            Input #1, classList(classCtr), classLevel(classCtr)
        Loop
    Close #1
End Sub
'Logout subroutine, for ease of use, it will be use by many logOut buttons in order to ensure that the proper varaibles are beign rest each time
Public Sub LogOut()
    Dim pos As Integer
    'Shows the login form
    frmPreLogin.Show
    'Tests to see if this is the first time the person has logged on to the system
    If firstTime And Not administrator And addedFlashVocab Then
        Open App.Path & "\Data\AddedFlashVocab.txt" For Append As #1
            Write #1, userName(StudentPosition) 'Writes the username of the student after logout in order noting that the user has logged in at leat once
        Close #1
    End If
    'permanantly records the score data, and keeps both login data and score data concurrent in permanent textfiles.
    Open App.Path & "\Data\Scores.txt" For Output As #2
    Open App.Path & "\Data\LoginData.txt" For Output As #3
        For pos = 1 To loginCtr
            Write #2, studentGradeName(pos), StudentGrade(pos), studentCorrect(pos), studentWrong(pos), StudentAttempted(pos)
            Write #3, userName(pos), firstName(pos), lastName(pos), password(pos), ClassEnrolled(pos)
        Next pos
    Close #3
    Close #2
    'Clears the student data to make sure that there are no errors when using the program
    StudentName = ""
    StudentPosition = 0
    
    
End Sub

'used to swap two strings in an array
Public Sub SwapString(ByRef String1 As String, ByRef String2 As String)
    Dim Temp As String
    Temp = String1
    String1 = String2
    String2 = Temp
End Sub
'Used to swap to singles in an array
Public Sub SwapSingle(ByRef Single1 As Single, ByRef Single2 As Single)
    Dim Temp As Single
    Temp = Single1
    Single1 = Single2
    Single2 = Temp
End Sub
'Use3d to Swap integers in an array
Public Sub SwapInteger(ByRef Integer1 As Integer, ByRef Integer2 As Integer)
    Dim Temp As Integer
    Temp = Integer1
    Integer1 = Integer2
    Integer2 = Temp
End Sub

Public Sub ReadNouns()
    'Public Subroutine to read the data from the nouns text file and store them in public variables
    NounCtr = 0
    Open App.Path & "\Data\Nouns.txt" For Input As #1
    
        Do Until EOF(1)
            NounCtr = NounCtr + 1
            Input #1, NomSNoun(NounCtr), GenSNoun(NounCtr), stemNoun(NounCtr), DeclensionNoun(NounCtr), GenderNoun(NounCtr), definitionNoun(NounCtr), NounDifficulty(NounCtr)
        Loop
    Close #1
End Sub

Public Sub ReadVerbs()
'Reads the text file for the list of verbs used in this program
    'Opens the text File
    Open App.Path & "\Data\Verbs.txt" For Input As #1
    'initializes the verbctr variable
    verbCtr = 0
    'Loops to read the file (until the end of file)
    Do Until EOF(1)
    'increments the verbCtr variable storing the size of the textfile/array
        verbCtr = verbCtr + 1
        'inputs the data into appropriate program level arrays
        Input #1, VerbPresStem(verbCtr), VerbInfinitive(verbCtr), VerbPerfStem(verbCtr), VerbPartStem(verbCtr), VerbDefinition(verbCtr), VerbConjugation(verbCtr), VerbDifficulty(verbCtr), VerbClass(verbCtr), VerbPrincipleParts(verbCtr)
    Loop
    Close #1
End Sub
'Used to get the class level for the current student in order to test the nouns and verbs agaisnt it
Public Sub GetClassLevel()
    'read Classes into appropriate arrays
    Call ReadClasses
    ' declares necessary variables for a match and stop search
    Dim foundClass As Boolean
    Dim classPos As Integer
    'Initializes varaibles
    foundClass = False
    classPos = 0
    'loops to search for the class which matches the class enrolled of the student position of the current user
    Do Until foundClass Or classPos = classCtr
        classPos = classPos + 1
        If ClassEnrolled(StudentPosition) = classList(classPos) Then
            foundClass = True
        End If
    Loop
    'Sets the student level public varaible for further use in the program
    If foundClass Then
        StudentLevel = classLevel(classPos)
    Else
        MsgBox "Class not found!?"
    End If
    
End Sub
'Used to calculate and update the various scores for the current student,
Public Sub CalculateGrade(ByRef correct As Integer, ByRef wrong As Integer, ByRef attempted As Integer)
    studentWrong(StudentPosition) = studentWrong(StudentPosition) + wrong
    studentCorrect(StudentPosition) = studentCorrect(StudentPosition) + correct
    StudentAttempted(StudentPosition) = StudentAttempted(StudentPosition) + attempted
    StudentGrade(StudentPosition) = (studentCorrect(StudentPosition) / StudentAttempted(StudentPosition))
End Sub

Public Sub MakeConcurrent()
    'Ensures that the data for these two text files are always concurrent after sorting and the like
    Dim pos As Integer
    
    Open App.Path & "\Data\Scores.txt" For Output As #2
    Open App.Path & "\Data\LoginData.txt" For Output As #3
        For pos = 1 To loginCtr
            Write #2, studentGradeName(pos), StudentGrade(pos), studentCorrect(pos), studentWrong(pos), StudentAttempted(pos)
            Write #3, userName(pos), firstName(pos), lastName(pos), password(pos), ClassEnrolled(pos)
        Next pos
    Close #3
    Close #2
End Sub
