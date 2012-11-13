VERSION 5.00
Begin VB.Form FrmProtein 
   BackColor       =   &H000000FF&
   Caption         =   "Protein"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12405
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   5400
      ScaleHeight     =   1875
      ScaleWidth      =   6555
      TabIndex        =   10
      Top             =   8760
      Width           =   6615
   End
   Begin VB.CommandButton cmdSearchCarbs 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Click here to search the Carb content of any of the foods listed above!"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8880
      Width           =   2895
   End
   Begin VB.CommandButton cmdSortProtein 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort list according to g of Protein (Most-Least)"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton cmdShowFood 
      BackColor       =   &H00FF0000&
      Caption         =   "Show list of Protein Rich Foods"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   3120
      ScaleHeight     =   5475
      ScaleWidth      =   8955
      TabIndex        =   3
      Top             =   1920
      Width           =   9015
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturnNutrition 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Nutrition"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturnMain 
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
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H000000FF&
      Caption         =   "For those who are Carb conscious:"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   8040
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "The average woman is      supposed to have          approximately 50-60     grams of protein per                 day!"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblProtein1 
      BackColor       =   &H000000FF&
      Caption         =   "  Protein"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "FrmProtein"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare global variables
Dim FoodType(1 To 100) As String
Dim CarbAmount(1 To 100) As Single
Dim ProteinAmount(1 To 100) As Integer
Dim Ctr As Integer
Dim J As Integer
Dim Pass As Integer
Dim Pos As Integer
Dim TempAmount As Integer
Dim TempFood As String
Dim TempCarb As Single
    'Bennie Health Project
    'FrmProtein
    'Heidi Donnelly
    'Written on: 9/29
    'The purpose of this form is to allow the user to find various foods that are rich in Protein as well as how much Protein is necessary. The user is also able to look into the amount of carbohydrates found in certain protein rich foods.
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturnMain_Click()
    FrmProtein.Hide
    FrmMain.Show
End Sub

Private Sub cmdReturnNutrition_Click()
    FrmProtein.Hide
    FrmNutritionMain.Show
End Sub
Private Sub cmdShowFood_Click()
    'this button will open, read, and put into parallel arrays a file that contains a list of protein-rich foods and then display it in a picture box
    
    'initialize variable(s)
    Ctr = 0
    
    'open file
    Open App.Path & "\ProteinFoods.txt" For Input As #1
    
    'read and place into arrays
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, FoodType(Ctr), ProteinAmount(Ctr), CarbAmount(Ctr)
    Loop
    
    'close file
    Close #1
    
    'print header and separator
    picResults.Print "   Food:"; Tab(16); "Approx. Amount of Protein (g):"
    picResults.Print "************************************************************"
    
    'print list in picture box
    For J = 1 To Ctr
        picResults.Print FoodType(J); Tab(25); ProteinAmount(J)
    Next J
End Sub

Private Sub cmdSortProtein_Click()
'this button will sort the list of protein rich foods according to the amount of protein found in them (most to least)
    
    'clear picture box
    picResults.Cls
    
    'sort list
    For Pass = 1 To Ctr + 1
        For Pos = 1 To Ctr - Pass
            If ProteinAmount(Pos) < ProteinAmount(Pos + 1) Then
                TempAmount = ProteinAmount(Pos)
                ProteinAmount(Pos) = ProteinAmount(Pos + 1)
                ProteinAmount(Pos + 1) = TempAmount
                TempFood = FoodType(Pos)
                FoodType(Pos) = FoodType(Pos + 1)
                FoodType(Pos + 1) = TempFood
                TempCarb = CarbAmount(Pos)
                CarbAmount(Pos) = CarbAmount(Pos + 1)
                CarbAmount(Pos + 1) = TempCarb
            End If
        Next Pos
    Next Pass
    
    'print header and separator
    picResults.Print "   Food:"; Tab(16); "Approx. Amount of Protein (g):"
    picResults.Print "*********************************************************"
    
    'print list in picture box
    For J = 1 To Ctr
        picResults.Print FoodType(J); Tab(25); ProteinAmount(J)
    Next J
End Sub
Private Sub cmdSearchCarbs_Click()
'this button will ask for a particular food and search for the amount of carbs found in that food

'declare variables
Dim Found As Boolean
Dim Food As String

'ask for foodtype
 Food = InputBox("Please enter the food you which to be searched. (Please type exactly as listed above)")
 
 'initialize variables
 J = 0
 Found = False
 
 'search until found or end of list
 Do While (Not Found) And J < Ctr
    J = J + 1
        If Food = FoodType(J) Then
            Found = True
        End If
Loop

If (Not Found) Then
    picResults2.Print Food; "is not in the list provided above."
Else
    picResults2.Print FoodType(J); " contain(s) approx. "; CarbAmount(J); "carbohydrates per serving."
End If
 
End Sub


