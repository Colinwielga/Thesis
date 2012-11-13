VERSION 5.00
Begin VB.Form frmWheel 
   BackColor       =   &H00000000&
   Caption         =   "Spin Time!"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   Picture         =   "frmWheel.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd5002 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmd600 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmd800 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmd700 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmd900 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   12
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmd1100 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmd1200 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmd400 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmd950 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmd750 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmd300 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmd1000 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmd450 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmd200 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmd500 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmd1300 
      Caption         =   "$$"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdChoose 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FIRST! Choose A Letter"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblPick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Now Pick a Number Sign!"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wheel of Fortune!(WheelofFortune.vbp)
'Form name: frmScreen2(WheelFortune.frm); Form caption: Puzzle
'Author: Maria Zipp
'Date written: 1st November, 2006
'Form Objective: This form is what determines how much money the
'               user gets per letter guessed. The user "spins the wheel"
'               (picks a spot on the wheel) and pushes the appropriate
'               button. The button disappears and reveals an amount underneath.
'               The player then selects the "choose a letter" button, which displays
'               an inputbox, and once the player enters a consonant, the program verifies
'               that it is a consonant and hides this form and shows the main form again.

Private Sub cmd1000_Click()
    'hides button and adds value hidden under button to total
    'if the letter is in the puzzle
    cmd1000.Visible = False
    total = total + 1000
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd1100_Click()
    cmd1100.Visible = False
    total = total + 1100
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd1200_Click()
    cmd1200.Visible = False
    total = total + 1200
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd1300_Click()
    cmd1300.Visible = False
    total = total + 1300
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd200_Click()
    cmd200.Visible = False
    total = total + 200
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd300_Click()
    cmd300.Visible = False
    total = total + 300
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
    
End Sub

Private Sub cmd400_Click()
    cmd400.Visible = False
    total = total + 400
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd450_Click()
    cmd450.Visible = False
    total = total + 450
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd500_Click()
    cmd500.Visible = False
    total = total + 500
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd5002_Click()
    cmd5002.Visible = False
    total = total + 500
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd600_Click()
    cmd600.Visible = False
    total = total + 600
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd700_Click()
    cmd700.Visible = False
    total = total + 700
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd750_Click()
    cmd750.Visible = False
    total = total + 750
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd800_Click()
    cmd800.Visible = False
    total = total + 800
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd900_Click()
    cmd900.Visible = False
    total = total + 900
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmd950_Click()
    cmd950.Visible = False
    total = total + 950
    frmWheel.Visible = False
    frmScreen2.Visible = True
    cmdChoose.Visible = True
End Sub

Private Sub cmdChoose_Click()
    found = False
    counter = 0
    'user enters a letter via inputbox
    cletter = InputBox("Choose a consonant (in Caps)", "Enter")
    'match-and-stop search to verify letter
    Do Until found = True Or counter > 21
        counter = counter + 1
        If alphaArray(counter) = cletter Then
            found = True
        Else
            found = False
        End If
    Loop
    If found = True Then
        position = InStr("ELMERFUDD", cletter)
        If position > 0 Then
            cmdChoose.Visible = False
        Else
            MsgBox "Sorry, your letter is not in the puzzle!", , "Try Again"
        End If
    ElseIf found = False Then
        MsgBox "Please enter your CONSONANT in CAPS", , "Error"
    End If
    
End Sub



