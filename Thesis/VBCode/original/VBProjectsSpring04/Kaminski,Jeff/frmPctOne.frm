VERSION 5.00
Begin VB.Form frmPctOne 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInput 
      Caption         =   "Manually Input Grades"
      Height          =   1335
      Left            =   2520
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "MS Notepad Text File"
      Height          =   1335
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "How would you like to submit your coursework percentage grades?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmPctOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

End Sub

Private Sub cmdInput_Click()

CtrPct = 0

Grade = InputBox("Please type your FIRST Percentage grade (0-100) ONLY and press OK.")

Do While Grade <> -999
    CtrPct = CtrPct + 1
    PctGrade(CtrPct) = Grade
    Grade = InputBox("Please type your Next Percentage grade (0-100) ONLY and press OK. Enter -999 if your are finished!")
Loop

frmPctOne.Hide
frmScale.Show

End Sub

Private Sub cmdText_Click()

CtrPct = 0

FileLocation = InputBox("Please type in the ENTIRE path for your Notepad file. Only Grades 0-100!")

Open FileLocation For Input As #1

Do While Not EOF(1)
    CtrPct = CtrPct + 1
    Input #1, PctGrade(CtrPct)
Loop
    
Close #1

frmPctOne.Hide
frmScale.Show

End Sub
