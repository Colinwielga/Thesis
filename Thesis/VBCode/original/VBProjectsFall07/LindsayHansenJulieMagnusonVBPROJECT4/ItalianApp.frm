VERSION 5.00
Begin VB.Form ItalianApp 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      Height          =   1455
      Left            =   480
      ScaleHeight     =   1395
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton cmdSeePicIA 
      BackColor       =   &H00FF8080&
      Caption         =   "Show Me the Picture please!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdShowDir 
      BackColor       =   &H00FF8080&
      Caption         =   "Show Me the Directions please!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowIn 
      BackColor       =   &H00FF8080&
      Caption         =   "Show Me the Ingredients please!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdSwitch4 
      BackColor       =   &H00FF8080&
      Caption         =   "Back to Main Form"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2415
   End
   Begin VB.PictureBox picItalianApp 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   3120
      ScaleHeight     =   3435
      ScaleWidth      =   7395
      TabIndex        =   2
      Top             =   1080
      Width           =   7455
   End
   Begin VB.CommandButton cmdSortAlph 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort Ingredients in Alphabetical Order (make sure ingredients is open)"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdSearchIn 
      BackColor       =   &H00FF8080&
      Caption         =   "Search for Ingredients"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label lblItalianAppTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Italian Grilled Bruschetta"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "ItalianApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer


Private Sub cmdSearchIn_Click()
Dim found As Boolean
Dim I As Integer
Dim NameOfIngredient As String

NameOfIngredient = InputBox("Enter an ingredient, make sure your spelling is correct!", "Ingredient Search")    'this creates an input box for user to input information
picItalianApp.Cls   'clears the form
I = 0
found = False

Do While ((Not found) And (I < ctr))    'this loop searches for the ingredient that the user is looking for
    I = I + 1
    If NameOfIngredient = ingr(I) Then found = True
Loop

If (Not found) Then 'this if statement prints out the ingredient or tells the user the ingredient is not in the recipe
    picItalianApp.Print NameOfIngredient; " was not in recipe!"
    Else
    picItalianApp.Print NameOfIngredient
End If

End Sub

Private Sub cmdSeePicIA_Click()
    Picture4.Picture = LoadPicture(App.Path & "\ItalianAppPic.jpg") 'this button when clicked loads a picture
End Sub

Private Sub cmdShowDir_Click()
Dim directions(1 To 20) As String

picItalianApp.Cls   'clears the form

Open App.Path & "\ItalianAppDirections.txt" For Input As #2 'this opens a file and inputs the directions in the form
    Do Until EOF(2)
    ctr = ctr + 1
        Input #2, directions(ctr)
        picItalianApp.Print directions(ctr)
    Loop
    Close #2

End Sub

Private Sub cmdShowIn_Click()

picItalianApp.Cls   'clears the form
ctr = 0

Open App.Path & "\ItalianAppIngredients.txt" For Input As #1    'this opens a file and inputs ingredients and amount in the form
    Do Until EOF(1)
    ctr = ctr + 1
        Input #1, amt(ctr), ingr(ctr)
        picItalianApp.Print amt(ctr), ingr(ctr)
    Loop
    Close #1
End Sub

Private Sub cmdSortAlph_Click()
Dim pass, temp2, I As Integer
Dim temp As String
Dim comp As Integer

picItalianApp.Cls   'clears the form

For pass = 1 To ctr - 1 'this puts the ingredients in alphabetical order
For comp = 1 To ctr - pass
    If ingr(comp) > ingr(comp + 1) Then
        temp = ingr(comp)
        temp2 = amt(comp)
        ingr(comp) = ingr(comp + 1)
        amt(comp) = amt(comp + 1)
        ingr(comp + 1) = temp
        amt(comp + 1) = temp2
    End If
Next comp
Next pass

For I = 1 To ctr
    picItalianApp.Print amt(I), ingr(I) 'this prints the amount and the ingredient
Next I


End Sub

Private Sub cmdSwitch4_Click()
Main.Show   'this shows the main form
ItalianApp.Hide 'this hides the italian appetizer form
End Sub
