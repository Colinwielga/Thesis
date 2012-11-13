VERSION 5.00
Begin VB.Form Snacks 
   BackColor       =   &H0080FFFF&
   Caption         =   "Snacks"
   ClientHeight    =   8100
   ClientLeft      =   5445
   ClientTop       =   2235
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   9555
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF80FF&
      Caption         =   "Back to Main Page"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdInputs 
      BackColor       =   &H00FF8080&
      Caption         =   "Input your snacks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton SnacksInputs 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Any snacks today?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   7695
      Left            =   5400
      ScaleHeight     =   7635
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   0
      Picture         =   "Snacks.frx":0000
      Top             =   240
      Width           =   4380
   End
   Begin VB.Label lblinstructions 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Snacks.frx":31D8E
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1920
      TabIndex        =   8
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label lblCitaion 
      Caption         =   "http://www.greetingbaskets.com/images/A2331RSRainbowSnacksFLg.jpg"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label lblInputs 
      BackColor       =   &H0080FFFF&
      Caption         =   "<======   Step #2"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblSnacks 
      BackColor       =   &H0080FFFF&
      Caption         =   "<=====   Step#1"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "Snacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Snacks Code
'Calorie Counter
'By: Ryan Sutherland and Tara Johnson
'3-26-08
'This button takes the user from the snakcs page to the main page.
Private Sub cmdBack_Click()
Snacks.Hide                         'Form switch
Main_Page.Show
End Sub
'Ends the program
Private Sub cmdQuit_Click()
End
End Sub
'This is a long Print statement. It was a different way to get the same result as the other forms.  These are the
'options that the user has when choosing what to input.
Private Sub SnacksInputs_Click()
picResults.Print Tab(0); "1"; Tab(15); "Potato Chips, 10 chips"
picResults.Print Tab(0); "2"; Tab(15); "Chocolate Chip Cookie"
picResults.Print Tab(0); "3"; Tab(15); "Cashews, 1 oz."
picResults.Print Tab(0); "4"; Tab(15); "Pistachios, 1 oz."
picResults.Print Tab(0); "5"; Tab(15); "Apple Sauce, 1 cup"
picResults.Print Tab(0); "6"; Tab(15); "Homemade Brownie"
picResults.Print Tab(0); "7"; Tab(15); "Cheese Crackers, 10 crackers"
picResults.Print Tab(0); "8"; Tab(15); "Ice Cream, 1 cup"
picResults.Print Tab(0); "9"; Tab(15); "Jelly Beans, 1 oz."
picResults.Print Tab(0); "10"; Tab(15); "Peach Slices, 1 cup"
picResults.Print Tab(0); "11"; Tab(15); "Popcorn(Unsalted), 1 cup"
picResults.Print Tab(0); "12"; Tab(15); "Pretzel Sticks, 10 Sticks"
picResults.Print Tab(0); "13"; Tab(15); "Choclolate Pudding, 1/2 cup"
picResults.Print Tab(0); "14"; Tab(15); "Vanilla Wafer Cookies, 10 cookies"
picResults.Print Tab(0); "15"; Tab(15); "Yogurt, 8 oz."
cmdInputs.Enabled = True                                            'Allows the user to use the input button.
End Sub
'This is an example of a Select Case. It basically is an if statement, but a more inclusive type. Here it does
'the same things as all the other forms, but is a different process.
Private Sub cmdInputs_Click()
Dim Choice As Single, CaloriesS As Single, Snacks As Integer
'An input box is still used so it appears exactly the same to the user.
Snacks = InputBox("Please enter the number of the corresponding food, then enter -999 when done. NOTE: Negative numbers that correspond to the choices will count toward the total.", "Dinner")
'This is a loop so that the user is allowed to enter many choices.  It will skip the loop if the input is -999, so
'that is why there is an input box before and at the end of this loop.
Do While Snacks <> -999
    Select Case Snacks
        Case 1
            CaloriesS = 105
            picResults.Print "Potato Chips, 10 chips"         'If #1 then add 105 and Print "Potato Chips, 10 chips".
        Case 2
            CaloriesS = 45
            picResults.Print "Chocolate Chip Cookie"
        Case 3
            CaloriesS = 165
            picResults.Print "Cashews, 1 oz."
        Case 4
            CaloriesS = 165
            picResults.Print "Pistachios, 1 oz."
        Case 5
            CaloriesS = 195
            picResults.Print "Apple Sauce, 1 cup"
        Case 6
            CaloriesS = 95
            picResults.Print "Homemade Brownie"
        Case 7
            CaloriesS = 50
            picResults.Print "Cheese Crackers, 10 crackers"
        Case 8
            CaloriesS = 270
            picResults.Print "Ice Cream, 1 cup"
        Case 9
            CaloriesS = 105
            picResults.Print "Jelly Beans, 1 oz."
        Case 10
            CaloriesS = 110
            picResults.Print "Peach Slices, 1 cup"
        Case 11
            CaloriesS = 30
            picResults.Print "Popcorn(Unsalted), 1 cup"
        Case 12
            CaloriesS = 10
            picResults.Print "Pretzel Sticks, 10 Sticks"
        Case 13
            CaloriesS = 155
            picResults.Print "Choclolate Pudding, 1/2 cup"
        Case 14
            CaloriesS = 185
            picResults.Print "Vanilla Wafer Cookies, 10 cookies"
        Case 15
            CaloriesS = 230
            picResults.Print "Yogurt, 8 oz."
        Case Else
                picResults.Print "This is not a snack choice."
    End Select
    'Another input box is needed to make sure the user continuosly enters numbers until -999.
    Snacks = InputBox("Please enter the number of the corresponding food, then enter -999 when done. NOTE: Negative numbers that correspond to the choices will count toward the total.", "Dinner")
    Calories_Snacks = Calories_Snacks + CaloriesS
Loop
'End of the loop and select case.
picResults.Print "You had"; Calories_Snacks; "Calories today outside of meal times."
End Sub
