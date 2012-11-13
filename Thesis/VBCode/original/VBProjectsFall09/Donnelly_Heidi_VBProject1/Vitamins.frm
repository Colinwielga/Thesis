VERSION 5.00
Begin VB.Form FrmVitamins 
   BackColor       =   &H0000FFFF&
   Caption         =   "Vitamins"
   ClientHeight    =   12180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13605
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12180
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   3720
      Picture         =   "Vitamins.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdVitD 
      BackColor       =   &H00C0C000&
      Caption         =   "Show list of foods that contain VITAMIN D"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9480
      Width           =   2175
   End
   Begin VB.CommandButton cmdVitC 
      BackColor       =   &H00C0C000&
      Caption         =   "Show list of foods that contain VITAMIN C"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdVitA 
      BackColor       =   &H00C0C000&
      Caption         =   "Show list of foods that contain VITAMIN A"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2175
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
      Top             =   120
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
      Top             =   120
      Width           =   1335
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1695
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
      Height          =   9375
      Left            =   3480
      ScaleHeight     =   9315
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   2520
      Width           =   9855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Vitamins!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1815
      Left            =   6360
      TabIndex        =   8
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "FrmVitamins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare global variables
Dim Ctr As Integer
Dim J As Integer
    'Bennie Health Project
    'FrmVitamins
    'Heidi Donnelly
    'Written on: 9/29
    'The purpose of this form is to allow the user to find various foods that are rich in three different vitamins.
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturnMain_Click()
    FrmVitamins.Hide
    FrmMain.Show
End Sub

Private Sub cmdReturnNutrition_Click()
    FrmVitamins.Hide
    FrmNutritionMain.Show
End Sub

Private Sub cmdVitA_Click()
 'this button will open, read, and put into an arrays a file that contains a list of Vitamin A-rich foods and then display it in a picture box
 
    'declare variables
    Dim Food(1 To 100) As String
    
    'initialize variable
    Ctr = 0
    
    'open file
    Open App.Path & "\VitaminA.txt" For Input As #1
    
    'read and place into arrays
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Food(Ctr)
    Loop
    
    'close file
    Close #1
    
    'clear picture box
    picResults.Cls
    
    'print intro and separator
    picResults.Print " The following foods are good sources of Vitamin A: "
    picResults.Print "***************************************************************"
    
    'print list in picture box
    For J = 1 To Ctr
        picResults.Print Food(J)
    Next J
    
End Sub

Private Sub cmdVitC_Click()
 'this button will open, read, and put into an arrays a file that contains a list of Vitamin C-rich foods and then display it in a picture box
 
    'declare variables
    Dim Food(1 To 100) As String
    
    'initialize variable
    Ctr = 0
    
    'open file
    Open App.Path & "\VitaminC.txt" For Input As #1
    
    'read and place into arrays
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Food(Ctr)
    Loop
    
    'close file
    Close #1
    
    'clear picture box
    picResults.Cls
    
    'print intro and separator
    picResults.Print " The following foods are good sources of Vitamin C: "
    picResults.Print "***************************************************************"
    
    'print list in picture box
    For J = 1 To Ctr
        picResults.Print Food(J)
    Next J
End Sub

Private Sub cmdVitD_Click()
    'this button displays a message regarding the sources of vitamin D using a picture box
    
    'clear picture box
    picResults.Cls
    
    'print message
    picResults.Print " ***Vitamin D is known as the sunshine vitamin since it is manufactured by the body "
    picResults.Print " after being exposed to sunshine. Ten to fifteen minutes of good sunshine three times"
    picResults.Print " weekly is adequate to produce the body's requirement of vitamin D. This means that we"
    picResults.Print " don't need to obtain vitamin D from our diet unless we get very little sunlight"
    picResults.Print " which is usually not a problem for most!"
End Sub

