VERSION 5.00
Begin VB.Form frmScale 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInputScale 
      BackColor       =   &H0080C0FF&
      Caption         =   "I would like to enter my own Grading Scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmd90Scale 
      BackColor       =   &H0080C0FF&
      Caption         =   ">=90 A     >=80 B     >=70 C     >=60 D    <60 F     "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Which Grading Scale would you like to use?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd90Scale_Click()

Path = "N:\CS130\handin\Kaminski, Jeff\"
CtrGScale = 0

Open Path & "scale.txt" For Input As #2

Do While Not EOF(2)
    CtrGScale = CtrGScale + 1
    Input #2, GradeScale(CtrGScale), GradeScaleLet(CtrGScale)
Loop

Close #2

frmScale.Hide
frmFinal.Show


End Sub

Private Sub cmdInputScale_Click()

CtrGScale = 0

GradeScaleTempLet = InputBox("Please type your FIRST LETTER GRADE of the grading scale ONLY and press OK.")
GradeScaleTemp = InputBox("Please type the LOWEST Percentage grade that recieves the letter grade just inputed (0-100) ONLY and press OK.")

Do While GradeScaleTemp <> -999
    CtrGScale = CtrGScale + 1
    GradeScale(CtrGScale) = GradeScaleTemp
    GradeScaleLet(CtrGScale) = GradeScaleTempLet
    GradeScaleTempLet = InputBox("Please type your NEXT LETTER GRADE of the grading scale ONLY and press OK. If you've entered your LAST Letter Grade already, type -999 and press OK")
    GradeScaleTemp = InputBox("Please type the LOWEST Percentage grade that recieves the letter grade just inputed (0-100) ONLY and press OK. If you've entered your LAST Percentage Grade already, type -999 and press OK")
Loop

frmScale.Hide
frmFinal.Show
End Sub
