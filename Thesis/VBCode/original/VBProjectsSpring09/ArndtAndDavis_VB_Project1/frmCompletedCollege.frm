VERSION 5.00
Begin VB.Form frmCompletedCollege 
   BackColor       =   &H000000C0&
   Caption         =   "Let's See What The Future Holds For Your Career!"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Submit Career and Continue"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   3735
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   11400
      TabIndex        =   7
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   12840
      TabIndex        =   6
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox txtCareer 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      TabIndex        =   5
      Top             =   7680
      Width           =   3375
   End
   Begin VB.CommandButton cmdTaxes 
      Caption         =   "Sort Careers In Descending Order Based On Amount Of Taxes Paid"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2760
      TabIndex        =   3
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdSalary 
      Caption         =   "Sort In Descending Order Careers Based On Salary"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2640
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
   End
   Begin VB.PictureBox picJobResults 
      Height          =   5775
      Left            =   5520
      ScaleHeight     =   5715
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   1560
      Width           =   7935
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "View List Of Possible Careers"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Careers That Do Not Require Degrees"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   9
      Top             =   480
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000040&
      FillColor       =   &H008080FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   735
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   8535
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Enter Profession you would choose as your ideal career"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      TabIndex        =   4
      Top             =   7680
      Width           =   3495
   End
End
Attribute VB_Name = "frmCompletedCollege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmCompletedCollege
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective: User views and sorts various possible career choices for someone who has completed some or all of college
'User can sort careers by income and taxes paid before choosing a career

Option Explicit

Dim JobsDegree(1 To 10) As String, Salary(1 To 10) As Double, Taxes(1 To 10) As Double, TempSalary As Double
Dim TempJobsDegree As String, TempTaxes As Double, Pass As Integer, Pos As Integer, I As Integer

Private Sub cmdList_Click()
'Read jobs with degree into array
picJobResults.Cls
ctr = 0


Open App.Path & "\JobsDegree.txt" For Input As #1
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, JobsDegree(ctr), Salary(ctr), Taxes(ctr)
Loop
Close #1

'print header
    picJobResults.Print "With a degree, you could have any of the following careers:"

'display all degree job options
For I = 1 To ctr
    picJobResults.Print JobsDegree(I)
Next I

End Sub

Private Sub cmdNext_Click()
'save users career choice for ending summary and continue to next form

Career = txtCareer.Text

frmTheFinerThingsInLife.Show
frmCompletedCollege.Hide

End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub cmdReturn_Click()
frmBeginning.Show
frmCompletedCollege.Hide

End Sub

Private Sub cmdSalary_Click()
'sort the careers in decending order based on salary using BUBBLE SORT

picJobResults.Cls

For Pass = 1 To ctr - 1
    For Pos = 1 To ctr - Pass
        If Salary(Pos) < Salary(Pos + 1) Then
            TempSalary = Salary(Pos)
            Salary(Pos) = Salary(Pos + 1)
            Salary(Pos + 1) = TempSalary
            TempJobsDegree = JobsDegree(Pos)
            JobsDegree(Pos) = JobsDegree(Pos + 1)
            JobsDegree(Pos + 1) = TempJobsDegree
            TempTaxes = Taxes(Pos)
            Taxes(Pos) = Taxes(Pos + 1)
            Taxes(Pos + 1) = TempTaxes
        End If
    Next Pos
Next Pass

For I = 1 To ctr
    picJobResults.Print "A " & JobsDegree(I) & " earns approximately " & FormatCurrency(Salary(I), 0) & " annually."
    Next I


End Sub

Private Sub cmdTaxes_Click()
'sort the careers in decending order based on taxes paid using BUBBLE SORT


picJobResults.Cls

For Pass = 1 To ctr - 1
    For Pos = 1 To ctr - Pass
        If Taxes(Pos) < Taxes(Pos + 1) Then
            TempSalary = Salary(Pos)
            Salary(Pos) = Salary(Pos + 1)
            Salary(Pos + 1) = TempSalary
            TempJobsDegree = JobsDegree(Pos)
            JobsDegree(Pos) = JobsDegree(Pos + 1)
            JobsDegree(Pos + 1) = TempJobsDegree
            TempTaxes = Taxes(Pos)
            Taxes(Pos) = Taxes(Pos + 1)
            Taxes(Pos + 1) = TempTaxes
        End If
    Next Pos
Next Pass

For I = 1 To ctr
    picJobResults.Print "A " & JobsDegree(I) & " must pay approximately " & FormatCurrency(Taxes(I), 0) & " annually."
    Next I
    
End Sub

