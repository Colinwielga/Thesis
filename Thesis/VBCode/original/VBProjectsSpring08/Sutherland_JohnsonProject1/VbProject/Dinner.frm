VERSION 5.00
Begin VB.Form Dinner 
   BackColor       =   &H000040C0&
   Caption         =   "Dinner"
   ClientHeight    =   9585
   ClientLeft      =   5445
   ClientTop       =   1845
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   9600
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdInputs 
      BackColor       =   &H008080FF&
      Caption         =   "Input your dinner choices"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   9015
      Left            =   5040
      ScaleHeight     =   8955
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton cmdDinnerInputs 
      BackColor       =   &H00FFC0C0&
      Caption         =   "What's For Dinner?"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack3 
      BackColor       =   &H0080FF80&
      Caption         =   "Back to main page"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label lblinstructions 
      BackColor       =   &H000080FF&
      Caption         =   $"Dinner.frx":0000
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   8
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label lblInputs 
      BackColor       =   &H000040C0&
      Caption         =   "<=====   Step#2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label lblDinner 
      BackColor       =   &H000040C0&
      Caption         =   "<=====    Step#1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblSitationOfPicture 
      Caption         =   "http://overthetop.beloblog.com/archives/2006/06/26/"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   3480
      Left            =   240
      Picture         =   "Dinner.frx":00A4
      Top             =   240
      Width           =   4620
   End
End
Attribute VB_Name = "Dinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dinner Code
'Calorie Counter
'By: Ryan Sutherland and Tara Johnson
'3-14-08
'This form has several different things going on. It will open a file, create arrays, exhaustively search the
'things told to search, contains a sentinel, a Do While loop, a few input boxes, and a math function.
Dim J As Integer, CTR As Integer, TotalCalories As Single
Dim DinnerNumber(1 To 28) As Integer, DinnerName(1 To 28) As String, DinnerCalories(1 To 28) As Single
'This brings the user from the dinner form to the starting form.
Private Sub cmdBack3_Click()
Dinner.Hide
Main_Page.Show
End Sub
'This button opens the notepad file and puts the information into arrays.
Private Sub cmdDinnerInputs_Click()
CTR = 0                                                 'CTR needs to be initially set to 0
J = 0
Open App.Path & "\Dinner.txt" For Input As #1           'Opens the file
    Do While Not EOF(1) And CTR < 28
        CTR = CTR + 1
        'Puts the information in the notepad into arrays.
        Input #1, DinnerNumber(CTR), DinnerName(CTR), DinnerCalories(CTR)
    Loop
Close                                         'The file must be closed when it is done being used.
'This is an example of a for next loop.  It is used in this case to print all the inputs from 1 to 28 in both
'of the two arrays.  It will go through and print all of the all items asked to print.
For J = 1 To 28
    picResults.Print DinnerNumber(J), DinnerName(J)
Next J
picResults.Print "*************************************"
cmdInputs.Enabled = True                     'Allows the user to click on the input button.
End Sub
'This button contains a Sentinel and an exhaustive search.  It gets a number and checks to make sure that
'it isn't the sentinel, and if not to add the corresponding calories.
Private Sub cmdInputs_Click()
Dim Choice As Single
CaloriesD = 0
J = 0
Choice = InputBox("Please enter the number of the corresponding food, then enter -999 when done. NOTE: Negative numbers that correspond to the choices will count toward the total.", "Dinner")
'This is where the example of an exhaustive search that has a sentinel in it start. We used them
'so that the program knows when the user wants it to stop searching and move onto the next process
'in the code.
Do While Choice <> -999
    For J = 1 To 28
        If Abs(Choice) = DinnerNumber(J) Then               'This ABS() will make any number typed a positive so that
            CaloriesD = DinnerCalories(J) + CaloriesD       'if a negative number is typed then it will still work.
            picResults.Print Left(DinnerName(J), 20)
        End If
    Next J
    'Here is an example of an input box that needs to be here so the user can input another choice.
    Choice = InputBox("Please enter the number of the corresponding food, then enter -999 when done. NOTE: Negative numbers that correspond to the choices will count toward the total.")
Loop
picResults.Print "*********************************************"
picResults.Print "Your total Dinner calories were"; CaloriesD       'Prints the calories
End Sub
'Ends the program
Private Sub cmdQuit_Click()
End
End Sub
