VERSION 5.00
Begin VB.Form nutritionform 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Nutrition form"
   ClientHeight    =   10065
   ClientLeft      =   4095
   ClientTop       =   780
   ClientWidth     =   7605
   LinkTopic       =   "Form2"
   ScaleHeight     =   10065
   ScaleWidth      =   7605
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFC0FF&
      Height          =   1335
      Left            =   1200
      Picture         =   "nutritionform.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   23
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   22
      Top             =   8760
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFC0FF&
      Height          =   975
      Left            =   5400
      Picture         =   "nutritionform.frx":0E1E
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   21
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFC0FF&
      Height          =   1095
      Left            =   3240
      Picture         =   "nutritionform.frx":16CB
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   20
      Top             =   3120
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFC0FF&
      Height          =   975
      Left            =   1080
      Picture         =   "nutritionform.frx":2262
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   19
      Top             =   3120
      Width           =   975
   End
   Begin VB.OptionButton opt3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Option3"
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Option2"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Option1"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   1755
      Picture         =   "nutritionform.frx":2CBD
      ScaleHeight     =   435
      ScaleWidth      =   4035
      TabIndex        =   10
      Top             =   240
      Width           =   4095
   End
   Begin VB.PictureBox picresults3 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   4320
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   9
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox txtamount 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdconvert 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to first form"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Nicole Diah  CS130"
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Strawberry"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Chocolate"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Vanilla"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0FF&
      Caption         =   "(enter a number 1-4 here for the serving amount )"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   $"nutritionform.frx":318B
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   11
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0FF&
      Caption         =   "4.  one half gallon"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "3.  one quart"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "2.  one pint"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "1.  one serving size (1/2 cup)"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "2) Select a serving amount: "
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "1) Choose a flavor:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "nutritionform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name : benjerryproject.vbp (Nicole Diah's VB Project.vbp)
'Form Name : nutritionform (nutritionform.frm)
'Author: Nicole Diah
'Date Written: Oct. 27, 2003
'Purpose of Form: Used to calculate the amount of sugar
                ' calories and fat in certain amounts of three flavors
                ' of Ben and Jerry's ice cream- vanilla, chocolate, strawberry
Dim J As Integer, Conversion(1 To 4) As Single, Path As String

Private Sub cmdback_Click()
Form1.Show
nutritionform.Hide
flavorsform.Hide
toursform.Hide

End Sub

Private Sub cmdconvert_Click()
Dim amount As Integer, flavor As Integer, sugar As Single, calories As Single
Dim fat As Single

picresults3.Cls 'clears the screen for new calculations
amount = txtamount

If opt1 = True Then
        sugar = 19 * Conversion(amount)
        calories = 240 * Conversion(amount)
        fat = 16 * Conversion(amount)
        'calculations if the flavor is vanilla
ElseIf opt2 = True Then
        sugar = 23 * Conversion(amount)
        calories = 270 * Conversion(amount)
        fat = 17 * Conversion(amount)
        'calculations if the flavor is chocolate
ElseIf opt3 = True Then
        sugar = 19 * Conversion(amount)
        calories = 180 * Conversion(amount)
        fat = 9 * Conversion(amount)
        'calculations if the flavor is strawberry
End If
picresults3.Print "Sugar:"; Tab(15); FormatNumber(sugar, 3); "g"
picresults3.Print "Calories:"; Tab(15); FormatNumber(calories, 3)
picresults3.Print "Total fat:"; Tab(15); FormatNumber(fat, 3); "g"


End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\Diah_Nicole\"
cmdconvert.Enabled = False
Open Path & "conversions.txt" For Input As #1
For J = 1 To 4
    Input #1, Conversion(J) 'input the conversion values
Next J
Close #1
End Sub

Private Sub txtamount_Change()
cmdconvert.Enabled = True
End Sub
