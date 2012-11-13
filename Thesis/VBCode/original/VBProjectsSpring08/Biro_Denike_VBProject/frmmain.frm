VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdRecipes 
      BackColor       =   &H0000C000&
      Caption         =   "Step 3: Recipes!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H0000FFFF&
      Caption         =   "Step 2: Plan a menu!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton cmdintake 
      BackColor       =   &H000080FF&
      Caption         =   "Step 1:What is my daily recommended caloric intake?"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.PictureBox picBackground 
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Bon Appetit:Menu Planner
'Form name: Main (frmMain.frm)
'Authors: Sarah Biro and Heather Denike
'Date written: 13 March 2008
'Purpose: This program allows users to calculate their
'recommended daily caloric intake range.
'It then offers a menu to choose items from and plan
'their own menu. It allows users to calculate the
'total amount of calories in their menu.
'The program allows users to see images and ingredients for select menu items.

Option Explicit

Private Sub cmdintake_Click()
    frmCaloricIntake.Show 'Brings user to the Caloric Intake form.
    frmmain.Hide
End Sub

Private Sub cmdMenu_Click()
    frmMenu.Show 'Brings user to the Menu form.
    frmmain.Hide
End Sub

Private Sub cmdQuit_Click()
    End 'Ends program.
End Sub

Private Sub cmdRecipes_Click()
    frmRecipes.Show 'Brings user to the Recipes form.
    frmmain.Hide
End Sub


Private Sub Form_Load()
    picBackground.Picture = LoadPicture(App.Path & "\produce.jpg") 'Loads image as background to the form.
End Sub

