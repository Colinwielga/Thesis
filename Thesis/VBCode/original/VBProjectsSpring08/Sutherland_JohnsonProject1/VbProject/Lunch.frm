VERSION 5.00
Begin VB.Form Lunch 
   BackColor       =   &H00008000&
   Caption         =   "Lunch"
   ClientHeight    =   8910
   ClientLeft      =   5040
   ClientTop       =   2040
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   10395
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdInputs 
      BackColor       =   &H008080FF&
      Caption         =   "Input your Lunch"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   8415
      Left            =   5760
      ScaleHeight     =   8355
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton cmdLunchInput 
      BackColor       =   &H0080FFFF&
      Caption         =   "What's For Lunch?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack2 
      BackColor       =   &H00FF8080&
      Caption         =   "Go back to main page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label lblinstructions 
      BackColor       =   &H0080FF80&
      Caption         =   $"Lunch.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2280
      TabIndex        =   8
      Top             =   6840
      Width           =   3255
   End
   Begin VB.Label lblInputs 
      BackColor       =   &H00008000&
      Caption         =   "<=====    Step#2"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblLunch 
      BackColor       =   &H00008000&
      Caption         =   "<=====    Step#1"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblSitationOfPicture 
      Caption         =   "http://www.jbsfamily.com/lunch.php"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   3480
      Left            =   480
      Picture         =   "Lunch.frx":00A4
      Top             =   360
      Width           =   4785
   End
End
Attribute VB_Name = "Lunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lunch Code
'Calorie Counter
'By: Ryan Sutherland and Tara Johnson
'3-15-08
'This form has several different things going on. It will open a file, create arrays, exhaustively search the
'things told to search, contains a sentinel, a Do While loop, a few input boxes, and a math function.
Dim J As Integer, CTR As Integer, TotalCalories As Single
Dim LunchNumber(1 To 22) As Integer, LunchName(1 To 22) As String, LunchCalories(1 To 22) As Single
'Takes the usr from the lunch form to the main page.
Private Sub cmdBack2_Click()
Lunch.Hide
Main_Page.Show
End Sub
'This is a button that reads a notepad file and puts that information into arrays.
Private Sub cmdLunchInput_Click()
CTR = 0
J = 0
Open App.Path & "\Lunch.txt" For Input As #1                'This is opening the file
    Do While Not EOF(1) And CTR < 22
        CTR = CTR + 1
        'These are the names of the arrays that the program put them into.
        Input #1, LunchNumber(CTR), LunchName(CTR), LunchCalories(CTR)
    Loop
Close                                                       'File must be closed.
'This is an example of a for next loop.  It is used in this case to print all the inputs from 1 to 22 in both
'of the two arrays.  It will go through and print all of the all items asked to print.
For J = 1 To 22
    picResults.Print LunchNumber(J), LunchName(J)
Next J
picResults.Print "*********************************************"
cmdInputs.Enabled = True                         'Makes the user able to acces the input button to type their choice.
End Sub
'This button contains a Sentinel and an exhaustive search.  It gets a number and checks to make sure that
'it isn't the sentinel, and if not to add the corresponding calories.
Private Sub cmdInputs_Click()
Dim Choice As Single
CaloriesL = 0
J = 0
Choice = InputBox("Please enter the number of the corresponding food, then enter -999 when done. NOTE: Negative numbers that correspond to the choices will count toward the total.", "Lunch")
'This is where the example of an exhaustive search that has a sentinel in it start. We used them
'so that the program knows when the user wants it to stop searching and move onto the next process
'in the code.
Do While Choice <> -999
    For J = 1 To 22
        If Abs(Choice) = LunchNumber(J) Then                 'Abs() takes the absolute value so negative numbers work.
            CaloriesL = LunchCalories(J) + CaloriesL
            picResults.Print Left(LunchName(J), 15)
        End If
    Next J
    'Here is an example of an input box that needs to be here so the user can input another choice.
    Choice = InputBox("Please enter the number of the corresponding food, then enter -999 when done. NOTE: Negative numbers that correspond to the choices will count toward the total.")
Loop
picResults.Print "*********************************************"
picResults.Print "Your total Lunch calories were"; CaloriesL            'Prints the calories
End Sub
'Ends the program
Private Sub cmdQuit_Click()
End
End Sub
