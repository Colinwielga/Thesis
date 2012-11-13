VERSION 5.00
Begin VB.Form FrmCalcium 
   BackColor       =   &H00C0C000&
   Caption         =   "Calcium"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12900
   FillColor       =   &H00C0C000&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowFood 
      BackColor       =   &H00FF0000&
      Caption         =   "Show List of Calcium Rich Foods"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   600
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   2415
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
      Left            =   3600
      ScaleHeight     =   5475
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   1200
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "  calcium, Calcium, CALCIUM!!!"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblCalcium1 
      BackColor       =   &H00C0C000&
      Caption         =   "  The average woman is         supposed to have         approximately 1000-1200    mg of calcium per day! "
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   3135
   End
End
Attribute VB_Name = "FrmCalcium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare global variables
Dim CalciumFoods(1 To 15, 1 To 3) As String
Dim Row As Integer
Dim Column As Integer
Dim Ctr As Integer
Dim J As Integer
    'Bennie Health Project
    'FrmCalcium
    'Heidi Donnelly
    'Written on: 9/23
    'The purpose of this form is to allow the user to find various foods that are rich in calcium as well as how much calcium is necessary
    
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturnMain_Click()
    FrmCalcium.Hide
    FrmMain.Show
End Sub

Private Sub cmdReturnNutrition_Click()
    'message the user to tell them about vitamins and proteins too
    MsgBox ("Calcium is very important ") & UserName & (", but make sure to check out Protein and Vitamins too!")
    FrmCalcium.Hide
    FrmNutritionMain.Show
End Sub

Private Sub cmdShowFood_Click()
    'this button will open, read, and put into parallel arrays a file that contains a list of calcium-rich foods and then display it in a picture box
    
    'initialize variable(s)
    Ctr = 0
    
    'open file
    Open App.Path & "\CalciumFoods.txt" For Input As #1
    
    'read and place into arrays
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, CalciumFoods(Ctr, 1), CalciumFoods(Ctr, 2), CalciumFoods(Ctr, 3)
        
    Loop
    
    'close file
    Close #1
    
    
    'print header and separator
    picResults.Print "   Food:"; Tab(20); "Serving Size (cup):"; Tab(45); "Approx. Amount of Calcium (mg):"
    picResults.Print "**************************************************************************************************"
    
    'print table in picture box
    For Row = 1 To 15
        picResults.Print CalciumFoods(Row, 1); Tab(20); CalciumFoods(Row, 2); Tab(45); CalciumFoods(Row, 3) 'prints one row on a single line
    Next Row
End Sub

