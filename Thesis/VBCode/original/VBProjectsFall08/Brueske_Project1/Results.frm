VERSION 5.00
Begin VB.Form Results 
   BackColor       =   &H8000000D&
   Caption         =   "Results"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Operations"
      Height          =   1335
      Left            =   3960
      TabIndex        =   6
      Top             =   1800
      Width           =   7215
      Begin VB.CommandButton cmdHome 
         Caption         =   "Home"
         Height          =   615
         Left            =   5280
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdMean 
         Caption         =   "Mean"
         Height          =   615
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdMedian 
         Caption         =   "Median "
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Data"
      Height          =   1335
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   7215
      Begin VB.CommandButton cmdshow 
         Caption         =   "Show Results"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearchMajors 
         Caption         =   "Search Fields"
         Height          =   615
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdSortUp 
         Caption         =   "Sort Field Asecending"
         Height          =   615
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton SortFieldsDown 
         Caption         =   "Sort Field Descending"
         Height          =   615
         Left            =   5160
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H0080FFFF&
      Height          =   6375
      Left            =   1320
      ScaleHeight     =   6315
      ScaleWidth      =   12675
      TabIndex        =   0
      Top             =   3360
      Width           =   12735
   End
End
Attribute VB_Name = "Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Alienation and Social Distance Project
'Results Form
'Kevin Brueske
'Created Oct 26, 2008
'Objective
    'Give user ways to manipulate and compare the preloaded or user created data.
    'Included: Show results, search, sort, median, and mode
    


Private Sub cmdhome_Click()
   'Return to the home form
    Results.Hide
    Home.Show
End Sub

Private Sub cmdMedian_Click()
'Determines the mean score of the alienation results or the social distance results
Dim answer As Single
Dim median As Single
Dim ctr As Integer
'Asks user to select which case
answer = InputBox("Median of 1 (Alienation Scores) or 2 (Social Distance Scores)", "Median")
Select Case answer
    'Alienation median
    Case Is = 1
        'Median formula
        median = (usrnum + 1) / 2
        picoutput.Cls
        picoutput.Print "The median alienation score is "; alienationScore(median)
    'Social distance median
    Case Is = 2
         'Median formula
        median = (usrnum + 1) / 2
        picoutput.Cls
        picoutput.Print "The median social distance score is "; distanceScore(median)
End Select
End Sub

Private Sub cmdMean_Click()
'Determines the median score of the alienation results or the social distance results
Dim answer As Single
Dim mean As Single
Dim ctr As Integer
'Determine whether to find the median of alienation scores of social distance scores
answer = InputBox("Mean of 1 (Alienation Scores) or 2 (Social Distance Scores)", "Mean")
Select Case answer
    'Alienation Scores Mean
    Case Is = 1
        'Adds all the alienation scores together
        For ctr = 1 To usrnum
            mean = mean + alienationScore(ctr)
        Next ctr
        'Determines mean value
        mean = FormatNumber((mean / usrnum), 1)
        
        picoutput.Cls
        picoutput.Print "The alienation mean score is"; mean
    'Social Distance Scores Mean
    Case Is = 2
        'Adds all the social distances scores together
        For ctr = 1 To usrnum
            mean = mean + distanceScore(ctr)
        Next ctr
          'Determines mean value
        mean = FormatNumber((mean / usrnum), 1)
        
        picoutput.Cls
        picoutput.Print "The social distance mean score is"; mean
    End Select

End Sub

Private Sub cmdSearchMajors_Click()
'Search function which allows user to search any field for any data entry
    Dim ctr As Single
    Dim qrandom As String
    Dim search As Single
    Dim qrandomage As Single
    'User selects which field to search
    search = InputBox("Enter 1 (First Name), 2 (Last Name), 3 (Age), 4 (Major), 5 (Social Class), 6 (Religion), 7 (Alienation Score), 8 (Distance Score)", "Search")
    Select Case search
        'First name Search
        Case Is = 1
            'Enter name
            qrandom = InputBox("Enter the First Name to search for.", "Name")
                picoutput.Cls
                picoutput.Print Tab(0); "|First Name|";
                picoutput.Print Tab(20); "|Last Name|";
                picoutput.Print Tab(40); "|Age|";
                picoutput.Print Tab(60); "|Major|";
                picoutput.Print Tab(80); "|Social Class|";
                picoutput.Print Tab(100); "|Religion|";
                picoutput.Print Tab(120); "|Alienation|";
                picoutput.Print Tab(140); "|Social Distance|"
                For ctr = 1 To usrnum
                    'If the name entered equals a name in the database then it will print out the data associated with the name
                    If LCase(fname(ctr)) = LCase(qrandom) Then
                           picoutput.Print Tab(0); fname(ctr);
                            picoutput.Print Tab(20); lname(ctr);
                            picoutput.Print Tab(40); age(ctr);
                            picoutput.Print Tab(60); major(ctr);
                            picoutput.Print Tab(80); socialclass(ctr);
                            picoutput.Print Tab(100); religion(ctr);
                            picoutput.Print Tab(120); alienationScore(ctr);
                            picoutput.Print Tab(140); distanceScore(ctr)
                    End If
                Next ctr
        'Last name search
        Case Is = 2
            'Enter last name
             qrandom = InputBox("Enter the Last Name to search for.", "Name")
                picoutput.Cls
                picoutput.Print Tab(0); "|First Name|";
                picoutput.Print Tab(20); "|Last Name|";
                picoutput.Print Tab(40); "|Age|";
                picoutput.Print Tab(60); "|Major|";
                picoutput.Print Tab(80); "|Social Class|";
                picoutput.Print Tab(100); "|Religion|";
                picoutput.Print Tab(120); "|Alienation|";
                picoutput.Print Tab(140); "|Social Distance|"
                For ctr = 1 To usrnum
                    'If the name entered equals a name in the database then it will print out the data associated with the name
                    If LCase(lname(ctr)) = LCase(qrandom) Then
                           picoutput.Print Tab(0); fname(ctr);
                            picoutput.Print Tab(20); lname(ctr);
                            picoutput.Print Tab(40); age(ctr);
                            picoutput.Print Tab(60); major(ctr);
                            picoutput.Print Tab(80); socialclass(ctr);
                            picoutput.Print Tab(100); religion(ctr);
                            picoutput.Print Tab(120); alienationScore(ctr);
                            picoutput.Print Tab(140); distanceScore(ctr)
                    End If
                Next ctr
            'Age search
            Case Is = 3
                'Input age
                 qrandomage = InputBox("Enter the age to search for.", "Age")
                picoutput.Cls
                picoutput.Print Tab(0); "|First Name|";
                picoutput.Print Tab(20); "|Last Name|";
                picoutput.Print Tab(40); "|Age|";
                picoutput.Print Tab(60); "|Major|";
                picoutput.Print Tab(80); "|Social Class|";
                picoutput.Print Tab(100); "|Religion|";
                picoutput.Print Tab(120); "|Alienation|";
                picoutput.Print Tab(140); "|Social Distance|"
                For ctr = 1 To usrnum
                    'If the age entered equals a agein the database then it will print out the data associated with the age
                    If age(ctr) = qrandomage Then
                           picoutput.Print Tab(0); fname(ctr);
                            picoutput.Print Tab(20); lname(ctr);
                            picoutput.Print Tab(40); age(ctr);
                            picoutput.Print Tab(60); major(ctr);
                            picoutput.Print Tab(80); socialclass(ctr);
                            picoutput.Print Tab(100); religion(ctr);
                            picoutput.Print Tab(120); alienationScore(ctr);
                            picoutput.Print Tab(140); distanceScore(ctr)
                    End If
                Next ctr
            'Major Search
            Case Is = 4
                'Enter major
                qrandom = InputBox("Enter the major to search for.", "Major")
                
                picoutput.Cls
                picoutput.Print Tab(0); "|First Name|";
                picoutput.Print Tab(20); "|Last Name|";
                picoutput.Print Tab(40); "|Age|";
                picoutput.Print Tab(60); "|Major|";
                picoutput.Print Tab(80); "|Social Class|";
                picoutput.Print Tab(100); "|Religion|";
                picoutput.Print Tab(120); "|Alienation|";
                picoutput.Print Tab(140); "|Social Distance|"
                For ctr = 1 To usrnum
                    'If the major entered equals a major in the database then it will print out the data associated with the major
                    If LCase(major(ctr)) = LCase(qrandom) Then
                           picoutput.Print Tab(0); fname(ctr);
                            picoutput.Print Tab(20); lname(ctr);
                            picoutput.Print Tab(40); age(ctr);
                            picoutput.Print Tab(60); major(ctr);
                            picoutput.Print Tab(80); socialclass(ctr);
                            picoutput.Print Tab(100); religion(ctr);
                            picoutput.Print Tab(120); alienationScore(ctr);
                            picoutput.Print Tab(140); distanceScore(ctr)
                    End If
                Next ctr
            'Social Class search
            Case Is = 5
                'Enter social class
                qrandom = InputBox("Enter the Social Class to search for.", "Social Class")
                picoutput.Cls
                picoutput.Print Tab(0); "|First Name|";
                picoutput.Print Tab(20); "|Last Name|";
                picoutput.Print Tab(40); "|Age|";
                picoutput.Print Tab(60); "|Major|";
                picoutput.Print Tab(80); "|Social Class|";
                picoutput.Print Tab(100); "|Religion|";
                picoutput.Print Tab(120); "|Alienation|";
                picoutput.Print Tab(140); "|Social Distance|"
                For ctr = 1 To usrnum
                    'If the social class entered equals a social class in the database then it will print out the data associated with the social class
                    If LCase(socialclass(ctr)) = LCase(qrandom) Then
                           picoutput.Print Tab(0); fname(ctr);
                            picoutput.Print Tab(20); lname(ctr);
                            picoutput.Print Tab(40); age(ctr);
                            picoutput.Print Tab(60); major(ctr);
                            picoutput.Print Tab(80); socialclass(ctr);
                            picoutput.Print Tab(100); religion(ctr);
                            picoutput.Print Tab(120); alienationScore(ctr);
                            picoutput.Print Tab(140); distanceScore(ctr)
                    End If
                Next ctr
            'Religion search
            Case Is = 6
                'Enter religion
                qrandom = InputBox("Enter the Religion to search for.", "Religion")
                picoutput.Cls
                picoutput.Print Tab(0); "|First Name|";
                picoutput.Print Tab(20); "|Last Name|";
                picoutput.Print Tab(40); "|Age|";
                picoutput.Print Tab(60); "|Major|";
                picoutput.Print Tab(80); "|Social Class|";
                picoutput.Print Tab(100); "|Religion|";
                picoutput.Print Tab(120); "|Alienation|";
                picoutput.Print Tab(140); "|Social Distance|"
                For ctr = 1 To usrnum
                    ''If the religion entered equals a religion in the database then it will print out the data associated with the religion
                    If LCase(religion(ctr)) = LCase(qrandom) Then
                           picoutput.Print Tab(0); fname(ctr);
                            picoutput.Print Tab(20); lname(ctr);
                            picoutput.Print Tab(40); age(ctr);
                            picoutput.Print Tab(60); major(ctr);
                            picoutput.Print Tab(80); socialclass(ctr);
                            picoutput.Print Tab(100); religion(ctr);
                            picoutput.Print Tab(120); alienationScore(ctr);
                            picoutput.Print Tab(140); distanceScore(ctr)
                    End If
                Next ctr
            'Alienation Score search
            Case Is = 7
                 'Enter alienation score
                 qrandomage = InputBox("Enter the Alienation Score to search for.", "Alienation Score")
                picoutput.Cls
                picoutput.Print Tab(0); "|First Name|";
                picoutput.Print Tab(20); "|Last Name|";
                picoutput.Print Tab(40); "|Age|";
                picoutput.Print Tab(60); "|Major|";
                picoutput.Print Tab(80); "|Social Class|";
                picoutput.Print Tab(100); "|Religion|";
                picoutput.Print Tab(120); "|Alienation|";
                picoutput.Print Tab(140); "|Social Distance|"
                For ctr = 1 To usrnum
                    'If the score entered equals a score in the database then it will print out the data associated with the score
                    If alienationScore(ctr) = qrandomage Then
                           picoutput.Print Tab(0); fname(ctr);
                            picoutput.Print Tab(20); lname(ctr);
                            picoutput.Print Tab(40); age(ctr);
                            picoutput.Print Tab(60); major(ctr);
                            picoutput.Print Tab(80); socialclass(ctr);
                            picoutput.Print Tab(100); religion(ctr);
                            picoutput.Print Tab(120); alienationScore(ctr);
                            picoutput.Print Tab(140); distanceScore(ctr)
                    End If
                Next ctr
            'Social distance score
            Case Is = 8
                'enter social distance score
                 qrandomage = InputBox("Enter the Distance Score to search for.", "Distance Score")
                picoutput.Cls
                picoutput.Print Tab(0); "|First Name|";
                picoutput.Print Tab(20); "|Last Name|";
                picoutput.Print Tab(40); "|Age|";
                picoutput.Print Tab(60); "|Major|";
                picoutput.Print Tab(80); "|Social Class|";
                picoutput.Print Tab(100); "|Religion|";
                picoutput.Print Tab(120); "|Alienation|";
                picoutput.Print Tab(140); "|Social Distance|"
                For ctr = 1 To usrnum
                       'If the score entered equals a score in the database then it will print out the data associated with the score
                    If distanceScore(ctr) = qrandomage Then
                           picoutput.Print Tab(0); fname(ctr);
                            picoutput.Print Tab(20); lname(ctr);
                            picoutput.Print Tab(40); age(ctr);
                            picoutput.Print Tab(60); major(ctr);
                            picoutput.Print Tab(80); socialclass(ctr);
                            picoutput.Print Tab(100); religion(ctr);
                            picoutput.Print Tab(120); alienationScore(ctr);
                            picoutput.Print Tab(140); distanceScore(ctr)
                    End If
                Next ctr
End Select
End Sub

Private Sub cmdshow_Click()
Dim ctr As Single
'Prints data as currently sorted for the user to see

picoutput.Cls
picoutput.Print Tab(0); "|First Name|";
picoutput.Print Tab(20); "|Last Name|";
picoutput.Print Tab(40); "|Age|";
picoutput.Print Tab(60); "|Major|";
picoutput.Print Tab(80); "|Social Class|";
picoutput.Print Tab(100); "|Religion|";
picoutput.Print Tab(120); "|Alienation|";
picoutput.Print Tab(140); "|Social Distance|"

'Output previously loaded data
For ctr = 1 To usrnum
    picoutput.Print Tab(0); fname(ctr);
    picoutput.Print Tab(20); lname(ctr);
    picoutput.Print Tab(40); age(ctr);
    picoutput.Print Tab(60); major(ctr);
    picoutput.Print Tab(80); socialclass(ctr);
    picoutput.Print Tab(100); religion(ctr);
    picoutput.Print Tab(120); alienationScore(ctr);
    picoutput.Print Tab(140); distanceScore(ctr)
Next ctr
End Sub

Private Sub cmdSortUp_Click()
'Sorts data by user preferred type into ascending order
    Dim ctr As Single
    Dim search As Single
    Dim pass As Single
    Dim pos As Single
    Dim temp As String
    Dim tempage As Single
        'Sort Algorithm
            '1.Ask the user which data type he or she would like to sort
            '2.Once the data type has been selected, it's time to bubble sort the data in
            'ascending Order
            '3. Look at the first value of the data type. Is it larger than the next value?
            '4. If so, "swap" it with the value to the right and swap any other data types
            'associated with that value.
            '5. Jump back to step 3, always moving to the next value. Do this until you have
            'reached the end of the data type values.
            '6. Now, move back to step 3 but start at the beginning of the data values.
            'Complete steps 2-5 once again. Continue doing this until you've done it n - 1
            'times (n = the number of data type values).
            '7. You've completed the bubble sort. Now you can write out your finished
            'answer.
            

    search = InputBox("Enter 1 (First Name), 2 (Last Name), 3 (Age), 4 (Major), 5 (Social Class), 6 (Religion), 7 (Alienation Score), 8 (Distance Score)", "Search")
    Select Case search
        Case Is = 1
            'Bubble sort code
            For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If fname(pos) > fname(pos + 1) Then
                    'Arrange array data to stay with the item being sorted
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
            Next pos
        Next pass
        'Last name sort
        Case Is = 2
             For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If lname(pos) > lname(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
            Next pos
        Next pass
        'Age sort
        Case Is = 3
             For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If age(pos) > age(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
            Next pos
        Next pass
        'Major sort
        Case Is = 4
             For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If major(pos) > major(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        'Social class sort
        Case Is = 5
            For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If socialclass(pos) > socialclass(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        'Religion sort
        Case Is = 6
        For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If religion(pos) > religion(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        'Alienation score sort
        Case Is = 7
            For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If alienationScore(pos) > alienationScore(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        'Social distance score sort
        Case Is = 8
        For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If distanceScore(pos) > distanceScore(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        End Select
        
            picoutput.Cls
            picoutput.Print Tab(0); "|First Name|";
            picoutput.Print Tab(20); "|Last Name|";
            picoutput.Print Tab(40); "|Age|";
            picoutput.Print Tab(60); "|Major|";
            picoutput.Print Tab(80); "|Social Class|";
            picoutput.Print Tab(100); "|Religion|";
            picoutput.Print Tab(120); "|Alienation|";
            picoutput.Print Tab(140); "|Social Distance|"
        'Output sorted data
        For ctr = 1 To usrnum
                picoutput.Print Tab(0); fname(ctr);
                picoutput.Print Tab(20); lname(ctr);
                picoutput.Print Tab(40); age(ctr);
                picoutput.Print Tab(60); major(ctr);
                picoutput.Print Tab(80); socialclass(ctr);
                picoutput.Print Tab(100); religion(ctr);
                picoutput.Print Tab(120); alienationScore(ctr);
                picoutput.Print Tab(140); distanceScore(ctr)
        Next ctr
        
End Sub

Private Sub SortFieldsDown_Click()
'Sorts data by user's preferred type by descending order
    Dim ctr As Single
    Dim search As Single
    Dim pass As Single
    Dim pos As Single
    Dim temp As String
    Dim tempage As Single
    'Input the type to sort in descending order
    search = InputBox("Enter 1 (First Name), 2 (Last Name), 3 (Age), 4 (Major), 5 (Social Class), 6 (Religion), 7 (Alienation Score), 8 (Distance Score)", "Search")
    Select Case search
        Case Is = 1
            'Bubble sort code
            For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If fname(pos) < fname(pos + 1) Then
                    'Change all array data to stay in the same array number as the moving sorted item
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
            Next pos
        Next pass
        'Last name sort
        Case Is = 2
             For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If lname(pos) < lname(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
            Next pos
        Next pass
        'Age sort
        Case Is = 3
             For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If age(pos) < age(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
            Next pos
        Next pass
        'Major Sort
        Case Is = 4
             For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If major(pos) < major(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        'Social Class sort
        Case Is = 5
            For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If socialclass(pos) < socialclass(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        'Religion sort
        Case Is = 6
        For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If religion(pos) < religion(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        'Aliention score sort
        Case Is = 7
            For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If alienationScore(pos) < alienationScore(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        
        Case Is = 8
        'Social Distance score sort
        For pass = 1 To usrnum - 1
                For pos = 1 To usrnum - pass
                    If distanceScore(pos) < distanceScore(pos + 1) Then
                    temp = fname(pos)
                    fname(pos) = fname(pos + 1)
                    fname(pos + 1) = temp
                    
                    temp = lname(pos)
                    lname(pos) = lname(pos + 1)
                    lname(pos + 1) = temp
                    
                    tempage = age(pos)
                    age(pos) = age(pos + 1)
                    age(pos + 1) = tempage
                    
                    temp = major(pos)
                    major(pos) = major(pos + 1)
                    major(pos + 1) = temp
                    
                    temp = socialclass(pos)
                    socialclass(pos) = socialclass(pos + 1)
                    socialclass(pos + 1) = temp
                    
                    temp = religion(pos)
                    religion(pos) = religion(pos + 1)
                    religion(pos + 1) = temp
                    
                    tempage = alienationScore(pos)
                    alienationScore(pos) = alienationScore(pos + 1)
                    alienationScore(pos + 1) = tempage
                    
                    tempage = distanceScore(pos)
                    distanceScore(pos) = distanceScore(pos + 1)
                    distanceScore(pos + 1) = tempage
                End If
                
            Next pos
        Next pass
        End Select
        
            picoutput.Cls
            picoutput.Print Tab(0); "|First Name|";
            picoutput.Print Tab(20); "|Last Name|";
            picoutput.Print Tab(40); "|Age|";
            picoutput.Print Tab(60); "|Major|";
            picoutput.Print Tab(80); "|Social Class|";
            picoutput.Print Tab(100); "|Religion|";
            picoutput.Print Tab(120); "|Alienation|";
            picoutput.Print Tab(140); "|Social Distance|"
        'Print out sorted data
        For ctr = 1 To usrnum
                picoutput.Print Tab(0); fname(ctr);
                picoutput.Print Tab(20); lname(ctr);
                picoutput.Print Tab(40); age(ctr);
                picoutput.Print Tab(60); major(ctr);
                picoutput.Print Tab(80); socialclass(ctr);
                picoutput.Print Tab(100); religion(ctr);
                picoutput.Print Tab(120); alienationScore(ctr);
                picoutput.Print Tab(140); distanceScore(ctr)
        Next ctr
        
End Sub
