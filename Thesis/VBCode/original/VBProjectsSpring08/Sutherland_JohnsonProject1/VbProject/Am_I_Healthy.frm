VERSION 5.00
Begin VB.Form Am_I_Healthy 
   BackColor       =   &H00FFFF00&
   Caption         =   "Total Calories"
   ClientHeight    =   8070
   ClientLeft      =   255
   ClientTop       =   1050
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   6285
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdClearScreen 
      BackColor       =   &H00FF0000&
      Caption         =   "Clear Screen"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdBack4 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to main page"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1215
   End
   Begin VB.PictureBox PicResults 
      Height          =   7815
      Left            =   120
      ScaleHeight     =   7755
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdCalories 
      BackColor       =   &H00FF0000&
      Caption         =   "Total Calories"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBMI 
      BackColor       =   &H00FF0000&
      Caption         =   "BMI?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Am_I_Healthy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Am I Healthy Code
'Calorie Counter
'By: Ryan Sutherland and Tara Johnson
'3-14-08
'This is the page in which most of the variable that are in the module are actually seem.  This is where
'the all get printed together to make a very nice and neat project.
Private Sub cmdBack4_Click()
Main_Page.Show          'This is a button that allows the user to go back to the home screen.
Am_I_Healthy.Hide
End Sub
'This allows the user to go to a seperate form that calculates their Body Mass Index.  Further explanation into
'this process is located on that specific page.
Private Sub cmdBMI_Click()
BMI_Index.Show
Am_I_Healthy.Hide
End Sub
'Here is the code for printing a variety of things.  It will print the total calories for breakfast,
'lunch and dinner. Then it prints a few different spacer type designs, followed by the number of calories
'that were consumed during snacking periods.It also allows the user to see a breakdown of how may calories
'are taken in during snack time.
Private Sub cmdCalories_Click()
Dim TotalCalories As Single, DayCalories As Single
TotalCalories = CaloriesB + CaloriesL + CaloriesD   'Equation to calculate meal calories
DayCalories = Calories_Snacks + TotalCalories       'Equation to calculate daily calories
picResults.Print "Your Breakfast calories for today were"; Tab(40); CaloriesB
picResults.Print "Your Lunch calories for today were"; Tab(40); CaloriesL
picResults.Print "Your Dinner calories for today were"; Tab(40); CaloriesD
picResults.Print "****************************************"
picResults.Print "Your calories for today during meals was"; TotalCalories
picResults.Print "                                                 "     'Creates a blank like when printing
picResults.Print Tab(19); "llllllll"                                     '"Look-good" design to show where
picResults.Print Tab(19); "llllllll"                                     'the important information is.
picResults.Print Tab(19); "llllllll"
picResults.Print Tab(19); "llllllll"
picResults.Print Tab(10); "\\\\\\\\\\//////////"
picResults.Print Tab(12); "\\\\\\\\////////"
picResults.Print Tab(14); "\\\\\\//////"
picResults.Print Tab(16); "\\\\////"
picResults.Print Tab(18); "\\//"
picResults.Print Tab(19); "\/"
picResults.Print "                                                 "
picResults.Print "                                                 "
picResults.Print "YOUR TOTAL DAILY CALORIES ARE BELOW:"
picResults.Print "Your out of meal calories for the day was"; Calories_Snacks
picResults.Print "Your calories for today during meals was"; TotalCalories
picResults.Print "****************************************"                     'Spacer for better appearence.
picResults.Print "Throughout the day today you consumed"; DayCalories
picResults.Print "***************************************************"
picResults.Print "Click on the BMI index button to check if you are"
picResults.Print "a healthy weight."
End Sub
'This button clears the screen so you can have a blank picture box.
Private Sub cmdClearScreen_Click()
picResults.Cls
End Sub
'Ending the program is the function of this button.
Private Sub cmdQuit_Click()
End
End Sub
