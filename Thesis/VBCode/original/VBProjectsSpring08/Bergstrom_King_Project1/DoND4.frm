VERSION 5.00
Begin VB.Form History 
   Caption         =   "Form1"
   ClientHeight    =   12645
   ClientLeft      =   4995
   ClientTop       =   2580
   ClientWidth     =   19050
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "DoND4.frx":0000
   ScaleHeight     =   12645
   ScaleWidth      =   19050
   Begin VB.CommandButton cmdAlphabetizeWinners 
      BackColor       =   &H80000003&
      Caption         =   "Alphabetize the Winners"
      Height          =   1215
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Width           =   7695
   End
   Begin VB.CommandButton cmdFindMinnesotans 
      BackColor       =   &H80000013&
      Caption         =   "Find all the Winners from Minnesota"
      Height          =   1215
      Index           =   1
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdReadFile 
      BackColor       =   &H8000000D&
      Caption         =   "Read List of Past Winners in the U.S."
      Height          =   1275
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   7695
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H80000013&
      Height          =   9015
      Left            =   120
      ScaleHeight     =   8955
      ScaleWidth      =   6435
      TabIndex        =   14
      Top             =   1080
      Width           =   6495
   End
   Begin VB.CommandButton cmdFindOregon 
      BackColor       =   &H8000000C&
      Caption         =   "Find all the Winners from Oregon"
      Height          =   1215
      Index           =   0
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CommandButton cmdFindDifferentName 
      BackColor       =   &H8000000D&
      Caption         =   "Translate"
      Height          =   1575
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox txtCountryNumber 
      Height          =   375
      Left            =   12000
      TabIndex        =   3
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton cmdHistoryBlurb 
      BackColor       =   &H80000013&
      Caption         =   "Where did it all begin?"
      Height          =   1215
      Left            =   6840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   7695
   End
   Begin VB.CommandButton cmdBacktoStart 
      BackColor       =   &H8000000A&
      Caption         =   "Back to Main Menu"
      Height          =   1215
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Enter Number Here ====>"
      Height          =   255
      Left            =   9720
      TabIndex        =   12
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Type:"
      Height          =   255
      Left            =   9840
      TabIndex        =   11
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "5 for the Philippeans"
      Height          =   255
      Left            =   10320
      TabIndex        =   10
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "4 for Zimbabwe"
      Height          =   255
      Left            =   10320
      TabIndex        =   9
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "3 for Greece"
      Height          =   255
      Left            =   10320
      TabIndex        =   8
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "2 for Germany"
      Height          =   255
      Left            =   10320
      TabIndex        =   7
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "1 for Mexico"
      Height          =   255
      Left            =   10320
      TabIndex        =   6
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Type in a number to see the show's title in another country!"
      Height          =   255
      Left            =   9840
      TabIndex        =   5
      Top             =   5400
      Width           =   4455
   End
   Begin VB.Label lblHistory 
      Caption         =   "The History of Deal or No Deal"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14415
   End
End
Attribute VB_Name = "History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Deal or No Deal Introduction
'Form Name: Start
'Authors: Chris Bergstrom and Brady King
'Date Written: March 27th, 2008

'Objective of Form: To inform the user of the origins of the show,
'as well as show the previous winners of the show,
'and manipulate the names of the previous winners and the state they're from.
'The form also translates the show's title in a foreign language so the user can see
'what the title would be if they were to watch it in a different country.

Private Sub cmdAlphabetizeWinners_Click() 'This alphabetizes the names to make the information more user friendly.

Dim Winner(1 To 100) As String      'Initializing the variables.
Dim HomeState(1 To 100) As String
Dim Move As Integer
Dim POS As Integer
Dim Temp As String
Dim B As Integer
Dim CTR As Integer
Dim Found As Boolean

picResults3.Cls                     'Clears the picture box.


Open App.Path & "\DoND_Winners.txt" For Input As #2     'Opens up the text document and puts it in an array.

Do While Not EOF(2)                             'Instructs to pull names from the text document if the end of the file has not been reached.
CTR = CTR + 1           'Keeps track of the number of times the computer looks through the list.
Input #2, Winner(CTR), HomeState(CTR)           'Places the information in the array.
Loop                                            'Repeats if needed.
Close #2                'Closes the array.

picResults3.Cls
picResults3.Print "Winner"; Tab(20); "Home State"               'Displays a header in the picture box.
picResults3.Print "----------------------------------------------------------------------"

For Move = 1 To CTR - 1                     'Sets a condition.
    For POS = 1 To CTR - Move               'Sets another condition.
        If Winner(POS) > Winner(POS + 1) Then       'Enters into an IF statement to re-order the names
                                                     'of the winners in alphabetical order.
        
            Temp = Winner(POS)
            Winner(POS) = Winner(POS + 1)
            Winner(POS + 1) = Temp
            
            Temp = HomeState(POS)
            HomeState(POS) = HomeState(POS + 1)
            HomeState(POS + 1) = Temp
                                   
            End If                          'Closes the IF statement.
            
Next POS                'Moves to the next position. If there are no more, it procedes to the next step listed.

Next Move               'Moves to the next time through the condition(if necessary).

For B = 1 To CTR        'Displays each winner's name and the corresponding state in the picture box.
picResults3.Print Winner(B); Tab(30); HomeState(B)

Next B                  'Moves to the next winner's name.

End Sub

Private Sub cmdBacktoStart_Click() 'Goes back to the greeting form.
History.Hide    'Hides the History form.
Start.Show      'Displays the Greeting form.
End Sub

Private Sub cmdFindDifferentName_Click() 'Translates the title of the show into a different language.
Dim Name As String          'Initializes the variables needed.
Dim CountryNumber As Integer

CountryNumber = txtCountryNumber        'Sets the variable to whatever is inputted into the text box.

    If CountryNumber = 1 Then                       'Runs the entered number through an IF statement
    MsgBox ("The title in Mexico is: Vas o No Vas.") 'to determine which category to place it and
    ElseIf CountryNumber = 2 Then                    'displays the corresponding text in a message box.
    MsgBox ("The title in Germany is: Die Show der GlucksSpirale.")
    ElseIf CountryNumber = 3 Then
    MsgBox ("The title in Greece is: Super Deal.")
    ElseIf CountryNumber = 4 Then
    MsgBox ("The title in Zimbabwe is: Saka Kana Aa Saka.")
    ElseIf CountryNumber = 5 Then
    MsgBox ("The title in the Philippeans is: Kapamilya.")
    Else
    MsgBox ("Sorry, that number is invalid.")
    
    End If          'Closes the IF statement.
    

End Sub

Private Sub cmdFindMinnesotans_Click(Index As Integer) 'This simply states that there are no winners from MN.
MsgBox ("There have been no winners from Minnesota. The first could be you!") 'Displays the appropriate text
End Sub                                                                       'in a message box.


Private Sub cmdFindOregon_Click(Index As Integer) 'This lists the number of winners from the state of Oregon.
Dim OREGON_PPL As Integer                           'This is used to display an array function.
Dim k As Integer
Dim Winner(1 To 100) As String                      'Initializes the variables needed.
Dim HomeState(1 To 100) As String

Open App.Path & "\DoND_States.txt" For Input As #3      'Opens up the text document and puts it in an array.

Do While Not EOF(3)

CTR = CTR + 1                                           'Keeps track of the number of times the computer looks through the list.

Input #3, HomeState(CTR), Winner(CTR)

Loop                                                    'Repeats if needed.

Close #3

OREGON_PPL = 0                                          'Sets the counter = 0.
Found = False                                           'Dictates that what is being looked for has not been found.
For k = 1 To CTR                                        'Enters into a FOR/NEXT loop.

If Len(HomeState(k)) = 6 Then                                           'Analyzes the length of the state names.
    MsgBox ("The following winners are from Oregon: " & Winner(k))      'Displays all the people from Oregon in a message box.
    OREGON_PPL = OREGON_PPL + 1                                         'Adds one to the counter.
    
    Found = True                                                        'The name(s) have been found.
    
    End If                                                              'Ends the IF statement.
    
Next k                                                                  'Closes the FOR/NEXT loop.

End Sub

Private Sub cmdHistoryBlurb_Click() 'Give a brief sentence on the date of the show's start.
MsgBox ("Deal or No Deal originally comes from the Netherlands where it debuted there in 2001. Since then more than 35 versions of the show have debuted from the United States to Indonesia.")
End Sub

Private Sub cmdReadFile_Click() 'Reads the winner's names into an array and displays them.
Dim Winner(1 To 100) As String              'Initializes the variables.
Dim HomeState(1 To 100) As String
Dim CTR As Integer

picResults3.Cls         'Clears the picture box.

CTR = 0                 'Sets the counter to zero.

Open App.Path & "\DoND_Winners.txt" For Input As #1                  'Opens up the text document and puts it in an array.
picResults3.Print "Winner", "Home State"                             'Displays the appropriate header.
picResults3.Print "-----------------------------------------"

Do While Not EOF(1)                                                  'Instructs to pull names from the text document if the end of the file has not been reached.

CTR = CTR + 1                                                        'Keeps track of the number of times the computer looks through the list.

Input #1, Winner(CTR), HomeState(CTR)                                'Places the information in the array.
        picResults3.Print Winner(CTR); Tab(20); HomeState(CTR)
        
Loop                                                                 'Repeats if needed.

MsgBox ("There have been " & CTR & " Winners.")                      'Displays the number of winners.

End Sub

