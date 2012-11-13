VERSION 5.00
Begin VB.Form frmBar 
   Caption         =   "The Bar"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   Picture         =   "frmBar.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStore 
      Height          =   615
      Left            =   8280
      Picture         =   "frmBar.frx":9AAB
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Shop till you drop."
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdBeerReferences 
      Caption         =   "Beer References"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "Having trouble picking a beer, click here."
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdLiquorReferences 
      Caption         =   "Liquor References"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Impress your friends with your knowledge of liquor."
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Height          =   255
      Left            =   7920
      Picture         =   "frmBar.frx":9CB2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "You don't want to leave yet, do you?"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Tools of the Trade"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Click here to learn about the tools used in bartending."
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdRecipes 
      Caption         =   "Drink Recipes"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Click Here to explore different drink  possibilities."
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Fred and Mike's School of Bartending"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmBar(The Bar)
'By Fred Paul & Michael McKeever
'March 22,2006
'This form is the Home page.  It is used to navigate to the different
'forms used in our program.

'Declare Your Name As a String function for input/output
Dim YourName As String

Private Sub cmdBeerReferences_Click()
'When clicked show beer references form and hide the bar form.
    frmBar.Hide
    frmBeerReferences.Show
End Sub

Private Sub cmdExit_Click()
   'The Exit Button will Display a thank you message to the customer displaying
   'Value for Your Name.
   'It will then ask the user for a yes/no input to leave the bar.
   
   MsgBox "Thank you for visiting Fred and Mike's School of Bartending!" & YourName, , "Come Again"
   Dim X As String
   X = InputBox("Are you sure you want to leave the Bar?")
    If X = "yes" Then
        End
    Else
        X = "no"
        frmBar.Show
    End If
        
    
    
   
      
    
End Sub

Private Sub cmdLiquorReferences_Click()
'When clicked Liquor references will Show the Liquor References form
'and hide the Bar form.
    frmBar.Hide
    frmLiquorReferences.Show
End Sub



Private Sub cmdRecipes_Click()
'When clicked Recipes will Show the Bartender form
'and hide the Bar form.
    frmBar.Hide
    frmBartender.Show
End Sub

Private Sub cmdStore_Click()
'When clicked The store will Show the Store form
'and hide the Bar form.
    frmBar.Hide
    frmStore.Show
End Sub

Private Sub cmdTools_Click()
'When clicked  Tolls of the Trade will show the BarTools form
'and hide the Bar form.
    frmBar.Hide
    frmBarTools.Show
End Sub

Private Sub Form_Load()
   'An Input Box will appear when program is stared asking the user for a name
   'This will be used to greet the customer on departing as well.
   
    YourName = InputBox("Input Your Name!", "Customer Bar Welcome")
End Sub
