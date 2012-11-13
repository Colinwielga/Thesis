VERSION 5.00
Begin VB.Form psychinfoform 
   BackColor       =   &H00C00000&
   Caption         =   "PSYCHOLOGY TERMS AND DEFINITIONS by SARAH AHLFS"
   ClientHeight    =   7485
   ClientLeft      =   5400
   ClientTop       =   3615
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdreturn1 
      BackColor       =   &H00808000&
      Caption         =   "RETURN TO ENTRY FORM"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton cmdpic10 
      BackColor       =   &H0000FFFF&
      Caption         =   "PIC 10"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdpic9 
      BackColor       =   &H008080FF&
      Caption         =   "PIC 9"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdpic8 
      BackColor       =   &H00FF00FF&
      Caption         =   "PIC 8"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdpic7 
      BackColor       =   &H00C0C000&
      Caption         =   "PIC 7"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdpic6 
      BackColor       =   &H000080FF&
      Caption         =   "PIC 6"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdpic5 
      BackColor       =   &H00FF0000&
      Caption         =   "PIC 5"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdpic4 
      BackColor       =   &H00C000C0&
      Caption         =   "PIC 4"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdpic3 
      BackColor       =   &H000000FF&
      Caption         =   "PIC 3"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdpic2 
      BackColor       =   &H0000C000&
      Caption         =   "PIC 2"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdpic1 
      BackColor       =   &H0080FFFF&
      Caption         =   "PIC 1"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox picturedisplaybox 
      BackColor       =   &H00FFFF80&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CLEAR OUTPUT BOX"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdsortalph 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SORT TERMS ALPHABETICALLY AND PRINT WITH DEFINITIONS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdadddata 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ADD PSYCH TERM AND DEFINITION TO FILE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdsearchdata 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SEARCH FOR PSYCH TERM DEFINITION"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdloaddata 
      BackColor       =   &H00C0FFC0&
      Caption         =   "LOAD DATA"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FF8080&
      FillColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9975
      Left            =   2520
      ScaleHeight     =   9915
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   480
      Width           =   12375
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   -240
         TabIndex        =   7
         Top             =   14520
         Width           =   16575
      End
   End
   Begin VB.Label lblpsychpics 
      BackColor       =   &H00000080&
      Caption         =   "SIGNIFICANT PEOPLE AND STUDIES IN PSYCHOLOGY"
      BeginProperty Font 
         Name            =   "MS PMincho"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   9600
      Width           =   2535
   End
   Begin VB.Label lblpsychprogram 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PSYCHOLOGY TERMS AND DEFINITIONS PROGRAM"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "psychinfoform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Psychology Terms and Definitions (psych_project.vbp)
'Psychinfoform (under psych_form.frm and psych_project.vbp)
'Psychentryform (under psychform2.frm and psych_project.vbp)
'Sarah Ahlfs
'October 20th, 2003
'Purpose: The overall purpose of this project is to have a database of psychology terms and definitions (since I'm a psych major) that can be easily seen, alphabetized, and searched through.  The pictures are reminders of some of the most significant people in psychology and 2 of my favorite psychology studies that have been done.
         'I will be able to use this throughout my career as a psychology major.  I can add to it and continually use it for quick reference and help on assignments if needed.
         'The purpose of this specific form (psychinfoform) is to contain all the information and ways it can be processed
Option Explicit
'these variables are globally dimensioned so that all subroutines can use them
Dim Psychterms(1 To 300) As String, Psychdefs(1 To 300) As String, ctr As Integer, pos As Integer
Dim enteredterm1 As String, entereddef1 As String
Public PATH As String
 

Private Sub cmdadddata_Click()
Dim no As String
no = "no" 'this initializes the the typed word no as the literal string word "no"
Open PATH & "psych.txt" For Append As #1 'this opened the file to write to it
        enteredterm1 = InputBox("Please enter the psychology term you wish to add to the file Or type no or hit cancel if you decide not to enter a term", "Enter a psychology term") 'this obtains the psych term to write to the file and gives user the option to decide not to enter a term
    If enteredterm1 = no Or enteredterm1 = "" Then 'this is a way to end the process if the user indicates that he/she no longer wants to enter data by typing the word no or hitting the cancel button
        Close #1 'this closes the file that was open to write to
    Else 'if the user enters a word other than no then the program moves to this next step
        entereddef1 = InputBox("Please enter the corresponding definition for the psychology term", "Enter Psychology Definition") 'this obtains the psych def to write to the file
         Write #1, enteredterm1, entereddef1 'this writes the psych term and def to the file
        MsgBox "The psychology term and definition have been added to the file psych.txt", , "Information Added" 'this message pops up
        MsgBox "REMINDER: New information cannot be properly accessed until the load button has been pushed", , "Push Load Button" 'this reminds users that the load button needs to be pushed inorder to see that new info is in the file, the search for new info, and to alphabetize file with new info
        Close #1 'closes the file
    End If
End Sub

Private Sub cmdclear_Click() 'this will clear the output in the picture box when this button is pushed
picresults.Cls
End Sub

Private Sub cmdloaddata_Click()
ctr = 0
picresults.Cls 'this clears any previous data that is in the picture box
picresults.Print "Psychology Term"; Tab(26); "Definition" 'the words in quotes will be printed and Tab allows the words to be entered in certain print columns
picresults.Print "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" 'this prints a dividing line
Open PATH & "psych.txt" For Input As #1 'this opens the file
Do While Not EOF(1) 'this loop moves the file into an array and prints it
    ctr = ctr + 1
    Input #1, Psychterms(ctr), Psychdefs(ctr) 'these are the 2 arrays that the file is moved into
    picresults.Print Psychterms(ctr); Tab(26); Psychdefs(ctr); Tab(1) 'this prints the 2 arrays in the indicated print columns
Loop
   Close #1 'this closes the opened file
   'the below buttons were not able to be accessed until the load button was hit
   cmdadddata.Enabled = True 'this allows the add data button to be accessed
   cmdsearchdata.Enabled = True 'this allows the search data button to be accessed
   cmdsortalph.Enabled = True 'this allows the sort data button to be accessed
End Sub

Private Sub cmdpic1_Click(Index As Integer)
picturedisplaybox.Picture = LoadPicture(PATH & "Wilhelm_Wundt.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Wilhelm Wundt, the father of psychology", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic10_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "allport.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Gordon Allport", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic2_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "William_James.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of William James, the father of American psychology", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic3_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "freud.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Sigmund Freud", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic4_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "bobodoll.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Bandura's study of aggression - Bobo doll", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic5_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "clothmom.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Harlow's cloth and wire mother monkey experiment", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic6_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "maslow.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Abraham Maslow", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic7_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "pavlov.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Ivan Pavlov", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic8_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "bandura.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Albert Bandura", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdpic9_Click()
picturedisplaybox.Picture = LoadPicture(PATH & "harlow.jpg") 'this loads the indicated picture into the picture box
MsgBox "This is a picture of Harry Harlow", , "What Is This Picture?" 'this message pops up on the screen
End Sub

Private Sub cmdquit_Click() 'when this button is hit it will end the running program
End
End Sub

Private Sub cmdreturn1_Click() 'when this button is hit the following things happen
psychinfoform.Visible = False 'psychinfoform-the form with all the info (this one) will disappear
psychentryform.Visible = True 'psychentryform -the entry form (other form) will show up
End Sub

Private Sub cmdsearchdata_Click()
Dim notfound As Boolean, termlook As String, no As String
no = "no" 'this initializes the the typed word no as the literal string word "no"
pos = 0 'this is another counter - it will keep the position
notfound = True 'this, while set at true, is saying that what you're searching for has not been found
Open PATH & "psych.txt" For Input As #1
MsgBox "REMINDER: enter word exactly as it is in file (with upper and lower case letters)or it won't be found", , "REMINDER" 'this message pops as a reminder to users
termlook = InputBox("Enter the psychology term you would like to look up", "Search for Psychology Term") 'the user enters the term he/she would like to look up in this box
'NOTE: there is no picresults.cls command in this button because I didn't want it to clear before every search incase the user wanted to keep adding to the list
Do While notfound And pos < ctr 'this searches through the array for the term entered in the input box (as long as position is less than ctr - so it doesn't go beyond the end of file) and prints it if it is found otherwise it moves to next step
    pos = pos + 1 'this increases the position by one each time through the loop
            If termlook = Psychterms(pos) Then 'this checks if the term typed in matches the psychterm at that position in the file
                picresults.Print termlook; Tab(26); Psychdefs(pos); Tab(1)
                notfound = False 'setting this variable to false now says that the you found what you were searching for
            End If
Loop
            If notfound Then 'if the term wasn't found then  the folling 2 messages pop up in boxes
                MsgBox "The psychology term you were looking for has not been found", , "Psycholgy Term Not Found"
                MsgBox "You will now have the opportunity to add the term and its definition if you would like to", , "Opportunity To Enter A Psychology Term and Definition"
                Close #1 'this closes file #1 for input
                Open PATH & "psych.txt" For Append As #1 'this opens file #1 to write to it
                enteredterm1 = InputBox("Please enter the psychology term you would like to add to the file OR type no or hit cancel if you don't wish to add a term", "Would You Like To Enter A Psychology Term?") 'this allows the user to enter a term to write to the file and allows them to not enter a term if they decide not to
                    If enteredterm1 = no Or enteredterm1 = "" Then 'if the person typed no or didn't type anything then the file will close and stop attempting to add data to the file
                        Close #1 'this closes the file
                Else 'if the person did type something other then no into the input box then it moves to the next step
                      entereddef1 = InputBox("Please enter the corresponding definition for the psychology term", "Enter Psychology Term Definition") 'this allows the user to enter the definition to write to the file
                        Write #1, enteredterm1, entereddef1 'this writes the entered term and definition to the file
                        MsgBox "The psychology term and defintion have been added to the file psych.txt", , "Information Added" 'this message pops up
                        MsgBox "REMINDER: New information cannot be properly accessed until the load button has been pushed", , "Push Load Button" 'this reminds users that the load button needs to be pushed inorder to see that new info is in the file, the search for new info, and to alphabetize file with new info
                    End If
         End If
Close #1 'this closes file #1 that was opened for writing to it
End Sub

Private Sub cmdsortalph_Click()
Dim pass As Integer, comp As Integer, tempterm As String, tempdef As String, J As Integer
Open PATH & "psych.txt" For Input As #1
picresults.Cls 'this clears any previous output in the picture box
picresults.Print "Psychology Term"; Tab(26); "Definition"
picresults.Print "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
'the below code will sort through the array and put the psych terms with their corresponding definitions into alphabetical order
For pass = 1 To ctr - 1 'this tells how many times to pass through the loop
    For comp = 1 To ctr - pass 'this tells how many comparisons to do
        If Psychterms(comp) > Psychterms(comp + 1) Then 'this and the next 3 equations alphabetizes the psych terms by comparing them to each other-one is always saved in a temporary holding spot so its not lost
            tempterm = Psychterms(comp)
            Psychterms(comp) = Psychterms(comp + 1)
            Psychterms(comp + 1) = tempterm
            tempdef = Psychdefs(comp) 'this and the next 2 equations makes sure that the psych definition prints with the proper psych term
            Psychdefs(comp) = Psychdefs(comp + 1)
            Psychdefs(comp + 1) = tempdef
        End If
    Next comp
Next pass
For J = 1 To ctr 'this goes through the now sorted array and prints the psych term and definition that is in the indicated position each time through the loop
    picresults.Print Psychterms(J); Tab(26); Psychdefs(J); Tab(1)
Next J
 Close #1 'this closes the file
End Sub

Private Sub Form_Load() 'everything under this subroutine is available to everything in this form
PATH = "N:\CS130\handin\Ahlfs_Sarah\" 'this allows me to not have to type the full open statement every time I want to open this directory path
End Sub

'NOTE: psychology defintions were taken from:Psychology Applied to Modern Life study guide and pictures were taken from http://psych.wisc.edu/henriques/resources/Images.html

