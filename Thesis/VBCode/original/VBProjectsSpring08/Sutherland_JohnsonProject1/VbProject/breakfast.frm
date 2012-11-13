VERSION 5.00
Begin VB.Form Breakfast 
   BackColor       =   &H000080FF&
   Caption         =   "Breakfast"
   ClientHeight    =   8895
   ClientLeft      =   5445
   ClientTop       =   1635
   ClientWidth     =   10005
   LinkTopic       =   "Form2"
   ScaleHeight     =   8895
   ScaleWidth      =   10005
   Begin VB.CommandButton cmdInputs 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to input your breakfast choices"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8160
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      Height          =   8535
      Left            =   6000
      ScaleHeight     =   8475
      ScaleWidth      =   3795
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdBreakfastInputs 
      BackColor       =   &H0000FFFF&
      Caption         =   "What's for Breakfast?"
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack1 
      Caption         =   "Back to main page"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label lblinstructions 
      BackColor       =   &H0080FFFF&
      Caption         =   $"breakfast.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2400
      TabIndex        =   8
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label lblInputs 
      BackColor       =   &H000080FF&
      Caption         =   "<=====   Step#2"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblBreafast 
      BackColor       =   &H000080FF&
      Caption         =   "<=====   Step #1"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblSitationOfPicture 
      Caption         =   "http://www.educationallearninggames.com/breakfast-foods-plastic-play-foods.asp"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   4275
      Left            =   120
      Picture         =   "breakfast.frx":00A4
      Top             =   120
      Width           =   4785
   End
End
Attribute VB_Name = "breakfast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Breakfast Code
'Calorie Counter
'By: Ryan Sutherland and Tara Johnson
'3-14-08
'This form has several different things going on. It will open a file, create arrays, exhaustively search the
'things told to search, contains a sentinel, a Do While loop, a few input boxes, and a math function.
Dim J As Integer, CTR As Integer, TotalCalories As Single
Dim BreakfastNumber(1 To 27) As Integer, BreakfastName(1 To 27) As String, BreakfastCalories(1 To 27) As Single
'Sends the user from the breakfast page to the Home page.
Private Sub cmdBack1_Click()
breakfast.Hide
Main_Page.Show
End Sub
'This is a button that first loads the file, then puts then into arrays, then prints the number the the user
'should type to get the calories that corresponds with that food name.
Private Sub cmdBreakfastInputs_Click()
CTR = 0                                                     'CTR needs to be initially set to 0
J = 0
Open App.Path & "\Breakfast.txt" For Input As #1            'Opening a file that contains data
    Do While Not EOF(1) And CTR < 27
        CTR = CTR + 1
        'Puts the data from the loaded file into three different arrays.
        Input #1, BreakfastNumber(CTR), BreakfastName(CTR), BreakfastCalories(CTR)
    Loop
Close                                                   'Closes the currently open file.
'This is an example of a for next loop.  It is used in this case to print all the inputs from 1 to 25 in both
'of the two arrays.  It will go through and print all of the all items asked to print.
For J = 1 To 27
    picResults.Print BreakfastNumber(J), BreakfastName(J)
Next J
picResults.Print "*************************************"
cmdInputs.Enabled = True                        'Allows the user to now click on what they want to input.
End Sub
'This button contains a Sentinel and an exhaustive search.  It gets a number and checks to make sure that
'it isn't the sentinel, and if not to add the corresponding calories.
Private Sub cmdInputs_Click()
Dim Choice As Single
CaloriesB = 0
J = 0
Choice = InputBox("Please enter the number of the corresponding food, then enter -999 when done. NOTE: Negative numbers that correspond to the choices will count toward the total.", "Breakfast")
'This is the example of an exhaustive search that has a sentinel in it so that it knows when the user wants it
'to stop searching and move onto the next process in the code.
Do While Choice <> -999
    For J = 1 To 27
        If Abs(Choice) = BreakfastNumber(J) Then             'This ABS() will make any number typed a positive so that
            CaloriesB = BreakfastCalories(J) + CaloriesB     'if a negative number is typed then it will still work.
            picResults.Print Left(BreakfastName(J), 15)
        End If
    Next J
    'Here is an example of an input box that needs to be here so the user can input another choice.
    Choice = InputBox("Please enter the number of the corresponding food, then enter -999 when done. NOTE: Negative numbers that correspond to the choices will count toward the total.")
'The Do while must loop back to the begining at some point so it does it here.
Loop
picResults.Print "*********************************************"
picResults.Print "Your total Breakfast calories were"; CaloriesB
End Sub
'Ends the project.
Private Sub cmdQuit_Click()
End
End Sub


