VERSION 5.00
Begin VB.Form frmRecipes 
   BackColor       =   &H0000C000&
   Caption         =   "Recipes"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picture1 
      BackColor       =   &H0000C000&
      Height          =   3255
      Left            =   3480
      ScaleHeight     =   3195
      ScaleWidth      =   3075
      TabIndex        =   11
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Main"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000C000&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   3075
      TabIndex        =   7
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton cmdTurkey 
      BackColor       =   &H000080FF&
      Caption         =   "Smoked Turkey Sandwich"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdrigatoni 
      BackColor       =   &H0000FFFF&
      Caption         =   "Rigatoni with Meat Sauce"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdTuna 
      BackColor       =   &H000080FF&
      Caption         =   "Tuna Fish Sandwich"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdsalmon 
      BackColor       =   &H0000FFFF&
      Caption         =   "Stuffed Broiled Salmon"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdburrito 
      BackColor       =   &H000080FF&
      Caption         =   "Breakfast Burrito"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdTaco 
      BackColor       =   &H0000FFFF&
      Caption         =   "Taco Salad"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblyum 
      BackColor       =   &H0000C000&
      Caption         =   "Yum!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblingredients 
      BackColor       =   &H0000C000&
      Caption         =   "Ingredients"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblClick 
      BackColor       =   &H0000C000&
      Caption         =   "Click an item to see the ingredients."
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmRecipes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Bon Appetit:Menu Planner
'Form name: Recipes (frmRecipes.frm)
'Authors: Sarah Biro and Heather Denike
'Date written: 3/13/2008
'Objective: This form allows users to click on select items from the menu.
            'When an item is clicked, the ingredients are displayed as well
            'as an image of the item.
Option Explicit

Private Sub cmdburrito_Click()
    picResults.Cls 'clears screen of any previous info

'prints all ingredients in breakfast burrito
    picResults.Print "1 flour tortilla"
    picResults.Print "(7'' diameter)"
    picResults.Print "1 scrambed egg"
    picResults.Print "(in 1 tsp soft margarine)"
    picResults.Print "1/3 cup black beans"
    picResults.Print "2 tbsp salsa"
    picture1.Cls
    picture1.Picture = LoadPicture(App.Path & "\burrito.jpg")
    'loads and displays item picture
End Sub



Private Sub cmdMain_Click()
'returns to main form
    frmmain.Show
    frmRecipes.Hide
End Sub

Private Sub cmdrigatoni_Click()
    picResults.Cls 'clears screen of any previous info

'prints all ingredients in rigatoni
    picResults.Print "1 cup rigatoni past"
    picResults.Print "(2 ounces dry)"
    picResults.Print "1/2 cup tomato sauce tomato bits"
    picResults.Print "2 ounces extra lean cooked ground beef"
    picResults.Print "(sauteed in 2 tsp vegetable oil)"
    picResults.Print "3 tbsp grated Parmesan cheese"
    picture1.Cls
    picture1.Picture = LoadPicture(App.Path & "\rigatoni.jpg")
    'loads and displays item picture
End Sub

Private Sub cmdsalmon_Click()
    picResults.Cls 'clears screen of any previous info

'prints all ingredients in salmon
    picResults.Print "5 ounce salmon filet"
    picResults.Print "1 ounce bread stuffing mix"
    picResults.Print "1 tbsp choppoed onions"
    picResults.Print "1 tbsp diced celery"
    picResults.Print "2 tsp canola oil"
    picture1.Cls
    picture1.Picture = LoadPicture(App.Path & "\salmon.jpg")
    'loads and displays item picture
End Sub

Private Sub cmdTaco_Click()
    picResults.Cls 'clears screen of any previous info

'prints all ingredients in taco
    picResults.Print "2 ounces tortilla chips"
    picResults.Print "2 ounces ground turkey, sauteed in"
    picResults.Print "2 tsp sunflower oil"
    picResults.Print "1/2 cup black beans"
    picResults.Print "1/2 cup iceberg lettuce"
    picResults.Print "2 slices tomato"
    picResults.Print "1 ounce low-fat cheddar cheese"
    picResults.Print " 2 tbsp salsa"
    picResults.Print "1/2 cup avocado"
    picResults.Print "1 tsp lime juice"
    picture1.Cls
    picture1.Picture = LoadPicture(App.Path & "\taco.jpg")
    'loads and displays item picture


End Sub

Private Sub cmdTuna_Click()
    picResults.Cls 'clears screen of any previous info

'prints all ingredients in tuna
    picResults.Print "2 slices rye bread"
    picResults.Print "3 ounces tuna"
    picResults.Print "(packed in water, drained)"
    picResults.Print "2 tsp mayonnaise"
    picResults.Print "1 tbsp diced celery"
    picResults.Print "1/4 cup shredded romaine lettuce"
    picResults.Print "2 slices tomato"
    picture1.Cls
    picture1.Picture = LoadPicture(App.Path & "\tuna.jpg")
    'loads and displays item picture
End Sub

Private Sub cmdTurkey_Click()
    picResults.Cls 'clears screen of any previous info

'prints all ingredients in turkey
    picResults.Print "2 ounces whole wheat pita bread"
    picResults.Print "1/4 cup romaine lettuce"
    picResults.Print "2 slices tomato"
    picResults.Print "3 ounces sliced smoked turkey breast"
    picResults.Print "1 tbsp mayo-type salad dressing"
    picResults.Print "1 tsp yellow mustard"
    picResults.Print
    picture1.Cls
    picture1.Picture = LoadPicture(App.Path & "\turkey.jpg")
    'loads and displays item picture
End Sub


