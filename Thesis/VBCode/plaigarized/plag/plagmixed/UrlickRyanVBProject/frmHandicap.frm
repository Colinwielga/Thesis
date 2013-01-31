VERSION 5.00
Begin VB.Form frmHandicap 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H00FF80FF&
      Caption         =   "Go Back to Main Menu"
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3600
      Width           =   2055
   End
   Begin VB.PictureBox picRedSlope 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   39
      Top             =   7680
      Width           =   1095
   End
   Begin VB.PictureBox picGoldSlope 
      BackColor       =   &H005ADAFA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   38
      Top             =   6840
      Width           =   1095
   End
   Begin VB.PictureBox picRedRating 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   37
      Top             =   7680
      Width           =   1095
   End
   Begin VB.PictureBox picGoldRating 
      BackColor       =   &H005ADAFA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   36
      Top             =   6840
      Width           =   1095
   End
   Begin VB.PictureBox picWhiteSlope 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   35
      Top             =   6000
      Width           =   1095
   End
   Begin VB.PictureBox picWhiteRating 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   34
      Top             =   6000
      Width           =   1095
   End
   Begin VB.PictureBox picBlueSlope 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   31
      Top             =   5160
      Width           =   1095
   End
   Begin VB.PictureBox picBlueRating 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   30
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdBlackberry 
      Height          =   735
      Left            =   120
      Picture         =   "frmHandicap.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5880
      Width           =   3135
   End
   Begin VB.CommandButton cmdTerritory 
      Height          =   1575
      Left            =   120
      Picture         =   "frmHandicap.frx":2579
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6720
      Width           =   3135
   End
   Begin VB.CommandButton cmdAlbany 
      Height          =   1095
      Left            =   120
      Picture         =   "frmHandicap.frx":51C7
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4680
      Width           =   3135
   End
   Begin VB.CommandButton cmdInputRating 
      BackColor       =   &H0080FF80&
      Caption         =   "Input Remaining"
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdStore 
      BackColor       =   &H008080FF&
      Caption         =   "Store Data"
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtFifthSlope 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtFourthSlope 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtThirdSlope 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtSecondSlope 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtFirstSlope 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtFifthRating 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtFourthRating 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtThirdRating 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtSecondRating 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtFirstRating 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox picFifthScore 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox picFourthScore 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox picThirdScore 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.PictureBox picSecondScore 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.PictureBox picFirstScore 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDisplaySlope 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Course Slope"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   33
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblDisplayRating 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Course Rating"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   32
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FF8080&
      Caption         =   "Don't know the rating or slope for a certain course? Click a button below to find out!"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Label lblInputAll 
      BackColor       =   &H00C0FFC0&
      Caption         =   "If the course rating and slope is the same for all scores, input rating and slope in first box and click ""Input Remaining"""
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00C0C0FF&
      Caption         =   $"frmHandicap.frx":61B1
      Height          =   1455
      Left            =   4800
      TabIndex        =   23
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblSlope 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Course Slope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   21
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblRating 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Course Rating"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblFifthScore 
      BackColor       =   &H000080FF&
      Caption         =   "Fifth score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblFourthScore 
      BackColor       =   &H0080C0FF&
      Caption         =   "Fourth Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblThirdScore 
      BackColor       =   &H000080FF&
      Caption         =   "Third score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblSecondScore 
      BackColor       =   &H0080C0FF&
      Caption         =   "Second score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblFirstScore 
      BackColor       =   &H000080FF&
      Caption         =   "First score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmHandicap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAlbany_Click()
    picBlueRating.Cls                   'clears picture boxes and displays data for the respective course
    picBlueSlope.Cls
    picWhiteRating.Cls
    picWhiteSlope.Cls








    picGoldRating.Cls
    picGoldSlope.Cls
    picRedRating.Cls   ' laksjd fsd fsdlfkj dslkf
    picRedSlope.Cls
    
    picBlueRating.Print "70.7"
    picBlueSlope.Print "122"
    picWhiteRating.Print "68.4"   ' laksjd fsd fsdlfkj dslkf
    picWhiteSlope.Print "117"
    picGoldRating.Print "64.1"
    picGoldSlope.Print "109"
    picRedRating.Print "65.8"   ' laksjd fsd fsdlfkj dslkf
    picRedSlope.Print "106"
End Sub

Private Sub cmdBlackberry_Click()
    picBlueRating.Cls
    picBlueSlope.Cls
    picWhiteRating.Cls
    picWhiteSlope.Cls













    picGoldRating.Cls
    picGoldSlope.Cls   ' laksjd fsd fsdlfkj dslkf
    picRedRating.Cls
    picRedSlope.Cls
    
    picBlueRating.Print "71.5"
    picBlueSlope.Print "127"
    picWhiteRating.Print "70.2"   ' laksjd fsd fsdlfkj dslkf
    picWhiteSlope.Print "121"
    picGoldRating.Print "67.7"
    picGoldSlope.Print "118"   ' laksjd fsd fsdlfkj dslkf
    picRedRating.Print "63.1"
    picRedSlope.Print "113"
End Sub

Private Sub cmdInputRating_Click()
    txtSecondRating.Text = txtFirstRating.Text          'fills in the rest of the text boxes with information from the first text box
    txtThirdRating.Text = txtFirstRating.Text








    txtFourthRating.Text = txtFirstRating.Text
    txtFifthRating.Text = txtFirstRating.Text
    
    txtSecondSlope.Text = txtFirstSlope.Text
    txtThirdSlope.Text = txtFirstSlope.Text
    txtFourthSlope.Text = txtFirstSlope.Text   ' laksjd fsd fsdlfkj dslkf
    txtFifthSlope.Text = txtFirstSlope.Text
End Sub

Private Sub cmdMenu_Click()
    frmHandicap.Hide
    frmStart.Show
End Sub

Private Sub cmdStore_Click()
    If Not (txtFirstRating.Text <> "" And txtSecondRating.Text <> "" And txtThirdRating.Text <> "" And txtFourthRating.Text <> "" And txtFifthRating.Text <> "" And txtFirstSlope.Text <> "" And txtSecondSlope.Text <> "" And txtThirdSlope.Text <> "" And txtFourthSlope.Text <> "" And txtFifthSlope.Text <> "") Then
        MsgBox "Please input a valid course rating for each of the 5 rounds.", , "Error"
    Else
        FirstRating = txtFirstRating.Text       'stores data into variables, computes the differential needed for the handicap
        SecondRating = txtSecondRating.Text     'computation, then stores differentials into an array
        ThirdRating = txtThirdRating.Text   ' laksjd fsd fsdlfkj dslkf









        FourthRating = txtFourthRating.Text
        FifthRating = txtFifthRating.Text
    
        FirstSlope = txtFirstSlope.Text
        SecondSlope = txtSecondSlope.Text   ' laksjd fsd fsdlfkj dslkf
        ThirdSlope = txtThirdSlope.Text









        FourthSlope = txtFourthSlope.Text   ' laksjd fsd fsdlfkj dslkf
        FifthSlope = txtFifthSlope.Text
    
        DiffArr(1) = (113 * 1 / 1 * (FirstScore - FirstRating)) / FirstSlope
        DiffArr(2) = (113 * 1 / 1 * (SecondScore - SecondRating)) / SecondSlope
        DiffArr(3) = (113 * 1 / 1 * (ThirdScore - ThirdRating)) / ThirdSlope
        DiffArr(4) = (113 * 1 / 1 * (FourthScore - FourthRating)) / FourthSlope
        DiffArr(5) = (113 * 1 / 1 * (FifthScore - FifthRating)) / FifthSlope
    
        Open App.Path & "\Differentials.txt" For Output As #1 'Writes data of array into a .txt file. Code obtained from Computing.net by Tom's Guide URL: http://www.computing.net/answers/programming/visual-basic-array/11213.html
        For Pos = 1 To 5
            Print #1, DiffArr(Pos)
        Next Pos
        Close #1
    
        frmHandicap.Hide            'returns to main menue form to compute handicap
        frmStart.Show
    
        txtFirstRating.Text = ""    'clears text boxes so the form can be used again
        txtSecondRating.Text = ""
        txtThirdRating.Text = ""
        txtFourthRating.Text = ""   ' laksjd fsd fsdlfkj dslkf
        txtFifthRating.Text = ""
    
        txtFirstSlope.Text = ""
        txtSecondSlope.Text = ""   ' laksjd fsd fsdlfkj dslkf
        txtThirdSlope.Text = ""
        txtFourthSlope.Text = ""
        txtFifthSlope.Text = ""
    End If
    
End Sub

Private Sub cmdTerritory_Click()
    picBlueRating.Cls
    picBlueSlope.Cls   ' laksjd fsd fsdlfkj dslkf
    picWhiteRating.Cls
    picWhiteSlope.Cls









    picGoldRating.Cls
    picGoldSlope.Cls   ' laksjd fsd fsdlfkj dslkf
    picRedRating.Cls
    picRedSlope.Cls
    
    picBlueRating.Print "72.9"
    picBlueSlope.Print "125"
    picWhiteRating.Print "71.0"   ' laksjd fsd fsdlfkj dslkf
    picWhiteSlope.Print "121"
    picGoldRating.Print "69.3"
    picGoldSlope.Print "116"
    picRedRating.Print "69.0"   ' laksjd fsd fsdlfkj dslkf
    picRedSlope.Print "120"
End Sub

Private Sub Form_Activate()
    picFirstScore.Cls                   'prints scores from input boxes that the user entered
    picSecondScore.Cls
    picThirdScore.Cls
    picFourthScore.Cls
    picFifthScore.Cls
    
    picFirstScore.Print FirstScore   ' laksjd fsd fsdlfkj dslkf
    picSecondScore.Print SecondScore   ' laksjd fsd fsdlfkj dslkf
    picThirdScore.Print ThirdScore
    picFourthScore.Print FourthScore
    picFifthScore.Print FifthScore
End Sub
