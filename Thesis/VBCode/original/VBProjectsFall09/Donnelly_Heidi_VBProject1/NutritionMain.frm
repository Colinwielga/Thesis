VERSION 5.00
Begin VB.Form FrmNutritionMain 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12195
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQ3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Question #3"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdQ2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Question #2"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnoughWater 
      BackColor       =   &H00FFFF00&
      Caption         =   "    <----------          Is that enough?"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturntoMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdVitamins 
      BackColor       =   &H000000FF&
      Caption         =   "Vitamins"
      DownPicture     =   "NutritionMain.frx":0000
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8400
      Picture         =   "NutritionMain.frx":1103
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton cmdProtein 
      BackColor       =   &H00004080&
      Caption         =   "Protein"
      DownPicture     =   "NutritionMain.frx":2206
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4800
      Picture         =   "NutritionMain.frx":279D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton cmdCalcium 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calcium"
      DownPicture     =   "NutritionMain.frx":2D34
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   840
      Picture         =   "NutritionMain.frx":3A9F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   3135
   End
   Begin VB.TextBox txtWaterConsumption 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdQ1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Question #1"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Test Your Nutrition Knowledge!!!"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   13
      Top             =   2040
      Width           =   7815
   End
   Begin VB.Label lblN3 
      BackColor       =   &H0000FF00&
      Caption         =   "**Find out foods that are good sources of:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   9
      Top             =   5400
      Width           =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   0
      X2              =   12120
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label lblN2 
      BackColor       =   &H0000FF00&
      Caption         =   "  How many glasses of water do you drink a day?"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   8175
   End
   Begin VB.Label lblN1 
      BackColor       =   &H0000FF00&
      Caption         =   "Nutrition"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   90
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2400
      TabIndex        =   0
      Top             =   -120
      Width           =   7455
   End
End
Attribute VB_Name = "FrmNutritionMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Bennie Health Project
    'FrmNutritionMain
    'Heidi Donnelly
    'Written: 9/23
    'The purpose of this form is to guide the user to three different necessary nutrient areas, it tests their knowledge about nutrition, and asks about their water consumption all while relaying information to better their nutrition.
    
Private Sub cmdEnoughWater_Click()
    'this button will calculate whether or not the amount of water the user entered into the text box is enough
    
    'declare variable
    Dim WaterConsumption As Integer
    
    'define variable
    WaterConsumption = txtWaterConsumption.Text
    
    'use if statement to assign a message box display
    If WaterConsumption >= 8 Then
        MsgBox ("That's awesome! It is important for you to continue to drink plenty of water because your whole system depends on it! Keep up the good work!")
    ElseIf WaterConsumption < 8 Then
        MsgBox ("Oh no! That's not good! It is recommended that you drink eight-eight ounce glasses of water a day in order to keep your body hydrated. Your body depends on water for many of its functions. It's important that you increase your water consumption in order to stay healthy!")
    Else
        MsgBox ("Please try again...invalid number!")
    End If

End Sub

Private Sub cmdProtein_Click()
    FrmNutritionMain.Hide
    FrmProtein.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturntoMain_Click()
    FrmNutritionMain.Hide
    FrmMain.Show
End Sub
Private Sub cmdQ1_Click()
'this button will ask the the user a trivia question #1 using an inputbox then display whether or not the answer is correct by using a message box

'declare variables
Dim Trivia1 As Integer

Trivia1 = InputBox("How many basic food groups are there? Note: Please answer using numbers.")
        If Trivia1 = 5 Then
        
            MsgBox ("That is correct! There are five basic food groups: (1)Dairy, (2)Meat/Eggs, (3)Grains/Bread, (4)Fruits, and (5)Vegetables.")
        Else
            MsgBox ("Sorry, that is incorrect! There are five basic food groups: (1)Milk Products, (2)Meat/Eggs, (3)Grains/Bread, (4)Fruits, and (5)Vegetables.")
        End If
End Sub
Private Sub cmdQ2_Click()
'this button will ask the the user a trivia question #2 using an inputbox then display whether or not the answer is correct by using a message box

'declare variables
Dim Trivia2 As Integer

Trivia2 = InputBox("Which of the five food groups provides a great source of calcium:  (1)Dairy  (2)Meat/Eggs  (3)Grains/Bread  (4)Fruits  (5)Vegetables?  Enter the number that corresponds with the food group.")
        If Trivia2 = 1 Then
            MsgBox ("Correct! Milk and other dairy products are a great source of calcium.")
        Else
            MsgBox ("Sorry, that is incorrect! The answer is (1)Dairy.")
        End If
End Sub
Private Sub cmdQ3_Click()
'this button will ask the the user a trivia question #3 using an inputbox then display whether or not the answer is correct by using a message box

'declare variables
Dim Trivia3 As Integer

Trivia3 = InputBox("What is the greatest source of Vitamin D? (1)Tomatoes (2)Whole Milk (3)Sunlight (4)Water (5)Red Meat Enter the number that corresponds with the correct answer. ")
        If Trivia3 = 3 Then
            MsgBox ("Correct! Being in the sun for just 10 minutes a day will provide you will all the Vitamin D you need!")
        Else
            MsgBox ("Sorry, that is incorrect! The SUN is the greatest source of Vitamin D")
        End If
End Sub

Private Sub cmdVitamins_Click()
    FrmNutritionMain.Hide
    FrmVitamins.Show
End Sub
Private Sub cmdCalcium_Click()
    FrmNutritionMain.Hide
    FrmCalcium.Show
End Sub
