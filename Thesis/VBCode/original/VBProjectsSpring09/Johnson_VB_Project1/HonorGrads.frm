VERSION 5.00
Begin VB.Form frmHonorGrads 
   BackColor       =   &H000080FF&
   Caption         =   "Honor Grads"
   ClientHeight    =   12645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18615
   LinkTopic       =   "Form1"
   ScaleHeight     =   12645
   ScaleWidth      =   18615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearchByHome 
      BackColor       =   &H8000000E&
      Caption         =   "Search By Home Town"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearchByLastName 
      BackColor       =   &H8000000E&
      Caption         =   "Search By Last Name"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdHargraveform 
      BackColor       =   &H8000000E&
      Caption         =   "Go back to Hargrave Main Page"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10080
      Width           =   1815
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H8000000E&
      Caption         =   "Sort by Last Names"
      Enabled         =   0   'False
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCollege 
      BackColor       =   &H8000000E&
      Caption         =   "Sort by Colleges"
      Enabled         =   0   'False
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdState 
      BackColor       =   &H8000000E&
      Caption         =   "Sort by States"
      Enabled         =   0   'False
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H8000000E&
      Caption         =   "Sort by Hometowns"
      Enabled         =   0   'False
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H8000000E&
      Caption         =   "Sort by First Names"
      Enabled         =   0   'False
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   12015
      Left            =   3240
      ScaleHeight     =   11955
      ScaleWidth      =   11475
      TabIndex        =   2
      Top             =   360
      Width           =   11535
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H8000000E&
      Caption         =   "Show"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000E&
      Caption         =   "Quit"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   11280
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Tim Johnson   3/20    To show some information about Hargraves Honor Grads and let users interact with that info"
      Height          =   615
      Left            =   15360
      TabIndex        =   11
      Top             =   11400
      Width           =   2895
   End
End
Attribute VB_Name = "frmHonorGrads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Last(1 To 150) As String, First(1 To 150) As String, Home(1 To 150) As String, State(1 To 150) As String, College(1 To 150) As String, Award(1 To 150) As String    'Declares all variables and strings for this form
Dim CTR As Integer, LookingFor As String, SearchOption As Integer, Found As Boolean, NumberFound As Integer
Dim pass As Integer, pos As Integer, j As Integer
Dim tempLast As String, tempFirst As String, tempHome As String, tempState As String, tempCollege As String, tempAward As String        'declares all variables and arrays for this form

Private Sub cmdCollege_Click()                          'sorts information by the college the graduate went on to attend

picResults.Cls                                          'clears the picture box

pass = 1
pos = 1
j = 0                                                   'preps the information to be sorted

For pass = 1 To CTR - 1                                 'sorts the information
    For pos = 1 To CTR - pass
        If College(pos) > College(pos + 1) Then
        
            tempLast = Last(pos)                       'sorts last names
            Last(pos) = Last(pos + 1)
            Last(pos + 1) = tempLast
            
            tempFirst = First(pos)                     'sorts first names
            First(pos) = First(pos + 1)
            First(pos + 1) = tempFirst
            
            tempHome = Home(pos)                       'sorts home towns of graduates
            Home(pos) = Home(pos + 1)
            Home(pos + 1) = tempHome
            
            tempState = State(pos)                     'sorts home states of graduates
            State(pos) = State(pos + 1)
            State(pos + 1) = tempState
            
            tempCollege = College(pos)                 'sorts colleges the grads went on to
            College(pos) = College(pos + 1)
            College(pos + 1) = tempCollege
            
            tempAward = Award(pos)                     'sorts the awards that graduates recieved
            Award(pos) = Award(pos + 1)
            Award(pos + 1) = tempAward
            
        End If
    Next pos
Next pass

    picResults.Print "Sorted by Colleges"              'preps a table to display information
    picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
    picResults.Print "***************************************************************************************************************************************************"

    For j = 1 To CTR                                    'actually prints the relevant information
             picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j)
    Next j
    
End Sub

Private Sub cmdFirst_Click()                          'sorts information by the college the graduate went on to attend

picResults.Cls                                        'clears the picture box

pass = 1
pos = 1
j = 0                                                 'preps the information to be sorted

For pass = 1 To CTR - 1                                'sorts the information
    For pos = 1 To CTR - pass
        If First(pos) > First(pos + 1) Then
        
            tempLast = Last(pos)                       'sorts last names
            Last(pos) = Last(pos + 1)
            Last(pos + 1) = tempLast
            
            tempFirst = First(pos)                     'sorts first names
            First(pos) = First(pos + 1)
            First(pos + 1) = tempFirst
            
            tempHome = Home(pos)                       'sorts home towns of graduates
            Home(pos) = Home(pos + 1)
            Home(pos + 1) = tempHome
            
            tempState = State(pos)                     'sorts home states of graduates
            State(pos) = State(pos + 1)
            State(pos + 1) = tempState
            
            tempCollege = College(pos)                 'sorts colleges the grads went on to
            College(pos) = College(pos + 1)
            College(pos + 1) = tempCollege
            
            tempAward = Award(pos)                     'sorts the awards that graduates recieved
            Award(pos) = Award(pos + 1)
            Award(pos + 1) = tempAward
            
        End If
    Next pos
Next pass

    picResults.Print "Sorted by First Names"              'preps a table to display information
    picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
    picResults.Print "***************************************************************************************************************************************************"

    For j = 1 To CTR                                    'actually prints the relevant information
             picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j)
    Next j

End Sub

Private Sub cmdHargraveForm_Click()                 'goes back to the Hargrave main form

frmHonorGrads.Hide                                  'goes back to the Hargrave main form
frmHargrave.Show

End Sub

Private Sub cmdHome_Click()                          'sorts information by the college the graduate went on to attend

picResults.Cls

pass = 1
pos = 1
j = 0                                                'preps the information to be sorted

For pass = 1 To CTR - 1                                'sorts the information
    For pos = 1 To CTR - pass
        If Home(pos) > Home(pos + 1) Then
        
            tempLast = Last(pos)                       'sorts last names
            Last(pos) = Last(pos + 1)
            Last(pos + 1) = tempLast
            
            tempFirst = First(pos)                     'sorts first names
            First(pos) = First(pos + 1)
            First(pos + 1) = tempFirst
            
            tempHome = Home(pos)                       'sorts home towns of graduates
            Home(pos) = Home(pos + 1)
            Home(pos + 1) = tempHome
            
            tempState = State(pos)                     'sorts home states of graduates
            State(pos) = State(pos + 1)
            State(pos + 1) = tempState
            
            tempCollege = College(pos)                 'sorts colleges the grads went on to
            College(pos) = College(pos + 1)
            College(pos + 1) = tempCollege
            
            tempAward = Award(pos)                     'sorts the awards that graduates recieved
            Award(pos) = Award(pos + 1)
            Award(pos + 1) = tempAward
            
        End If
    Next pos
Next pass

    picResults.Print "Sorted by Home Towns"              'preps a table to display information
    picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
    picResults.Print "***************************************************************************************************************************************************"

    For j = 1 To CTR                                    'actually prints the relevant information
             picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j)
    Next j
    
End Sub

Private Sub cmdLast_Click()                          'sorts information by the college the graduate went on to attend

picResults.Cls                                        'clears the picture box

pass = 1
pos = 1
j = 0                                                'preps the information to be sorted

For pass = 1 To CTR - 1                              'sorts the information
    For pos = 1 To CTR - pass
        If Last(pos) > Last(pos + 1) Then
        
            tempLast = Last(pos)                       'sorts last names
            Last(pos) = Last(pos + 1)
            Last(pos + 1) = tempLast
            
            tempFirst = First(pos)                     'sorts first names
            First(pos) = First(pos + 1)
            First(pos + 1) = tempFirst
            
            tempHome = Home(pos)                       'sorts home towns of graduates
            Home(pos) = Home(pos + 1)
            Home(pos + 1) = tempHome
            
            tempState = State(pos)                     'sorts home states of graduates
            State(pos) = State(pos + 1)
            State(pos + 1) = tempState
            
            tempCollege = College(pos)                 'sorts colleges the grads went on to
            College(pos) = College(pos + 1)
            College(pos + 1) = tempCollege
            
            tempAward = Award(pos)                     'sorts the awards that graduates recieved
            Award(pos) = Award(pos + 1)
            Award(pos + 1) = tempAward
            
        End If
    Next pos
Next pass

    picResults.Print "Sorted by Last Names"              'preps a table to display information
    picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
    picResults.Print "***************************************************************************************************************************************************"

    For j = 1 To CTR                                    'actually prints the relevant information
             picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j)
    Next j
    
End Sub

Private Sub cmdQuit_Click()             'ends the program
    End                                 'ends the program
End Sub

Private Sub cmdSearchByHome_Click()                          'searches by the home town the graduate is from

Found = False
j = 0
NumberFound = 0                                                                  'preps the information to be searched

LookingFor = InputBox("What Home Town are you looking for?", , "Enter a string")                        'asks the user what they are searching for

Do While SearchOption > 2 Or SearchOption < 1                                                           'forces the user to actually pick a specific type of search to do
SearchOption = InputBox("How do you want to Search? 1=Match/Stop, 2=Exhaustive", , "Enter a 1 or 2")    'lets the user pick what type of search to do
Loop

picResults.Cls                                                                  'clears the picture box

If SearchOption = 1 Then                                                        'if the user selected a match/stop search
    Do While Not Found And j <= CTR                                             'do until the first match is found
        j = j + 1
        If LookingFor = Home(j) Then                                           'if found then keep track of the finding
            Found = True
        End If
    Loop
    If Not Found Then
        picResults.Print "No one from "; LookingFor; " graduated from Hargrave in 2005." 'lets the user know the search failed
    Else
        picResults.Print "The first person from this home town is listed below"     'preps a table for information to be displayed
        picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
        picResults.Print "***************************************************************************************************************************************************"
        picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j) 'prints the match/stop search results
    End If
ElseIf SearchOption = 2 Then                                                    'if the person wants to do an exhaustive search
    For j = 1 To CTR
        If LookingFor = Home(j) Then                                            'if the search is a success
            Found = True                                                        'register that something has been found
            NumberFound = NumberFound + 1                                       'add more to the total found
            If NumberFound = 1 Then                                             'only lets the below table be printed once
                picResults.Print "The people from this home town are listed below"  'preps a table to display information
                picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
                picResults.Print "***************************************************************************************************************************************************"
            End If                                                  'prints all relevant information
            picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j)
        End If
    Next j
    If Not Found Then                                               'lets user know that the search they attempted, failed
        picResults.Print "No one from "; LookingFor; " graduated from Hargrave in 2005."
    End If                                                          'displays the total number of hits found on the exhaustive search
    picResults.Print "The number of people from this home town: "; NumberFound
End If


End Sub

Private Sub cmdSearchByLastName_Click()                                             'allows a person to either do a match/stop search or an exhaustive search by the last name of the graduate

Found = False
j = 0
NumberFound = 0                                                                  'preps the information to be searched

LookingFor = InputBox("Whose last name are you looking for?", , "Enter a string")                      'asks the user what they are searching for

Do While SearchOption > 2 Or SearchOption < 1                                                           'forces the user to actually pick a specific type of search to do
SearchOption = InputBox("How do you want to Search? 1=Match/Stop, 2=Exhaustive", , "Enter a 1 or 2")    'lets the user pick what type of search to do
Loop

picResults.Cls                                                                  'clears the picture box

If SearchOption = 1 Then                                                        'if the user selected a match/stop search
    Do While Not Found And j <= CTR                                             'do until the first match is found
        j = j + 1
        If LookingFor = Last(j) Then                                           'if found then keep track of the finding
            Found = True
        End If
    Loop
    If Not Found Then
        picResults.Print LookingFor; " did not graduate from Hargrave in 2005." 'lets the user know the search failed
    Else
        picResults.Print "The first person with this last name is listed below"     'preps a table for information to be displayed
        picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
        picResults.Print "***************************************************************************************************************************************************"
        picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j) 'prints the match/stop search results
    End If
ElseIf SearchOption = 2 Then                                                    'if the person wants to do an exhaustive search
    For j = 1 To CTR
        If LookingFor = Last(j) Then                                            'if the search is a success
            Found = True                                                        'register that something has been found
            NumberFound = NumberFound + 1                                       'add more to the total found
            If NumberFound = 1 Then                                             'only lets the below table be printed once
                picResults.Print "The people with this last name are listed below"  'preps a table to display information
                picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
                picResults.Print "***************************************************************************************************************************************************"
            End If                                                  'prints all relevant information
            picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j)
        End If
    Next j
    If Not Found Then                                               'lets user know that the search they attempted, failed
        picResults.Print "No one by the last name of "; LookingFor; " graduated from Hargrave in 2005."
    End If                                                          'displays the total number of hits found on the exhaustive search
    picResults.Print "The number of people with this last name: "; NumberFound
End If


End Sub

Private Sub cmdShow_Click()                                         'shows the information that was loaded below
Dim I As Integer                                                    'preps the information to be displayed
I = 0

    picResults.Cls                                                  'sets up a display table
    picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
    picResults.Print "***************************************************************************************************************************************************"

Do While I <= CTR                                                   'prints all relevant information
    I = I + 1
    picResults.Print Last(I); Tab(15); First(I); Tab(35); Home(I); Tab(55); State(I); Tab(65); College(I); Tab(105); Award(I)
Loop

cmdShow.Enabled = False                                             'allows other functions to be accessed
cmdFirst.Enabled = True
cmdLast.Enabled = True
cmdHome.Enabled = True
cmdState.Enabled = True
cmdCollege.Enabled = True


End Sub

Private Sub cmdState_Click()                          'sorts information by the college the graduate went on to attend

picResults.Cls                                        'clears the picture box

pass = 1
pos = 1
j = 0                                                 'preps the information to be sorted

For pass = 1 To CTR - 1                               'sorts the information
    For pos = 1 To CTR - pass
        If State(pos) > State(pos + 1) Then
        
            tempLast = Last(pos)                       'sorts last names
            Last(pos) = Last(pos + 1)
            Last(pos + 1) = tempLast
            
            tempFirst = First(pos)                     'sorts first names
            First(pos) = First(pos + 1)
            First(pos + 1) = tempFirst
            
            tempHome = Home(pos)                       'sorts home towns of graduates
            Home(pos) = Home(pos + 1)
            Home(pos + 1) = tempHome
            
            tempState = State(pos)                     'sorts home states of graduates
            State(pos) = State(pos + 1)
            State(pos + 1) = tempState
            
            tempCollege = College(pos)                 'sorts colleges the grads went on to
            College(pos) = College(pos + 1)
            College(pos + 1) = tempCollege
            
            tempAward = Award(pos)                     'sorts the awards that graduates recieved
            Award(pos) = Award(pos + 1)
            Award(pos + 1) = tempAward
            
        End If
    Next pos
Next pass

    picResults.Print "Sorted by States"                 'preps a new table
    picResults.Print "Last Name"; Tab(15); "First Name"; Tab(35); "Hometown"; Tab(55); "State"; Tab(65); "College"; Tab(105); "Award"
    picResults.Print "***************************************************************************************************************************************************"

    For j = 1 To CTR                                    'prints all relevant information
             picResults.Print Last(j); Tab(15); First(j); Tab(35); Home(j); Tab(55); State(j); Tab(65); College(j); Tab(105); Award(j)
    Next j
    
End Sub

Private Sub Form_Load()                                                                   'loads multiple arrays for this form
    CTR = 0

    Open App.Path & "\HonorGrads.txt" For Input As #1                                     'opens the text file to be read
    
    Do While Not EOF(1)                                                                   'reads all information into arrays until there is no more information
        
        CTR = CTR + 1
        
        Input #1, Last(CTR), First(CTR), Home(CTR), State(CTR), College(CTR), Award(CTR)
        
    Loop

    Close #1                                                                              'closes file after it is done being read

MsgBox "Your arrays have been loaded for this form", , "Loaded"                           'lets user know file was read


End Sub


