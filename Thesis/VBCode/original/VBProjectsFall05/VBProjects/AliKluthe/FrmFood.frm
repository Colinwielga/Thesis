VERSION 5.00
Begin VB.Form FrmFood 
   BackColor       =   &H0000C000&
   Caption         =   "Food"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Image ImgOils 
      Height          =   1605
      Left            =   6720
      Picture         =   "FrmFood.frx":0000
      Top             =   4200
      Width           =   1230
   End
   Begin VB.Image ImgDairy 
      Height          =   1575
      Left            =   7560
      Picture         =   "FrmFood.frx":058C
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Image ImgMeat 
      Height          =   1395
      Left            =   4200
      Picture         =   "FrmFood.frx":112B
      Top             =   4560
      Width           =   1440
   End
   Begin VB.Image ImgBread 
      Height          =   1170
      Left            =   3960
      Picture         =   "FrmFood.frx":1E6E
      Top             =   2400
      Width           =   1800
   End
   Begin VB.Image ImgVeggies 
      Height          =   1230
      Left            =   1200
      Picture         =   "FrmFood.frx":2710
      Top             =   4680
      Width           =   1830
   End
   Begin VB.Image ImgFruit 
      Height          =   1365
      Left            =   600
      Picture         =   "FrmFood.frx":3157
      Top             =   2280
      Width           =   1890
   End
   Begin VB.Label lblLearn 
      BackColor       =   &H0000C000&
      Caption         =   $"FrmFood.frx":40C6
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   8175
   End
   Begin VB.Label lblEatingRight 
      BackColor       =   &H0000C000&
      Caption         =   "Eating Right"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Be Healthy (VBFinalProject.vbp)
'Form Name: Main Menu (FrmMain.frm)
'Author: Ali Kluthe
'Date: 10/27/2005
'Objective: The purpose of this form is to teach the user about food. The user can find information about how many servings of each type of food they should eat each day.

Private Sub cmdReturn_Click() 'This button allows the user to return to the main menu.
FrmFood.Hide 'Hides the food form
FrmMain.Show 'Shows the main form
End Sub

Private Sub ImgBread_Click() 'This button allows the user to see a message box containing information about grain.
MsgBox "Each day you need to eat 6 ounces of grain per day. At least half of the grains you eat should be whole grain.", , "Grain" 'prints the words in quotations
End Sub

Private Sub ImgDairy_Click() 'This button allows the user to see a message box containing information about dairy.
MsgBox "Each day you should consume 3 cups of dairy. Choose low fat or fat free milk, yogurt or cheese.", , "Dairy" 'prints the word in quotations
End Sub

Private Sub ImgFruit_Click() 'This button allows the user to see a message box containing information about fruit.
MsgBox "Each day you need to eat 2 cups of fruits. The fruit can be fresh, frozen, canned or dried.", , "Fruit" 'prints the word in quotations
End Sub

Private Sub ImgMeat_Click() 'This button allows the user to see a message box containing information about meat.
MsgBox " Five and a half ounces of lean meat should be eaten each day. It is healthier if it is baked, broiled or grilled.Try to avoid fried meat. Fish, beans, peas and nuts are other good sources of protien.", , "Meat" 'prints the word in quotations
End Sub

Private Sub ImgOils_Click() 'This button allows the user to see a message box containing information about fats and oils.
MsgBox "Fats and Oils should be used sparingly. Read labels and avoid consuming food with trans fat in it. Trans fat increases your bad cholesterol and increases your risk for heart disease and diabetes.", , "Fats and Oils" 'prints the word in quotations
End Sub

Private Sub ImgVeggies_Click() 'This button allows the user to see a message box containing information about vegetables.
MsgBox "Two and a half cups of vegetables should be eaten every day. Try to eat a colorful variety of vegetables.", , "Vegetables" 'prints the word in quotations
End Sub

