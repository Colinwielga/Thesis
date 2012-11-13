VERSION 5.00
Begin VB.Form frmCognitiveDisorders 
   BackColor       =   &H0080FFFF&
   Caption         =   "Cognitive Disorders"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   ForeColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10320
      TabIndex        =   12
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdDefine 
      Caption         =   "Define Disorder!"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFFF&
      Height          =   3135
      Left            =   5040
      ScaleHeight     =   3075
      ScaleWidth      =   4995
      TabIndex        =   9
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   8
      Top             =   4800
      Width           =   3135
   End
   Begin VB.CommandButton cmdReturntoDisorders 
      Caption         =   "Return to Disorders"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturnHome 
      Caption         =   "Return Home"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdDefineCognitiveDisorders 
      Caption         =   "Define: Cognitive Disorders"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   11415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Input Disorder from above as shown:"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dementia"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   6
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Delirium"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   5
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Parkinson's Disorder"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cognitive Disorders"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "frmCognitiveDisorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    picResults.Cls
End Sub

Private Sub cmdDefine_Click()
Dim aaaa As Boolean
Dim bbbb As String
Dim cccc As Long
Dim dddd As Long
Dim eeee(1 To 100) As String
Dim ffff(1 To 100) As String
Dim gggg As Long
Dim hhhh As String

    Open App.Path & "\eeee.txt" For Input As #1
    dddd = -1 + 1 - 0
    bbbb = txtInput.Text
    Do While False = EOF(1)
        dddd = dddd + 1
        Input #1, eeee(dddd), gggg
        For cccc = 1 To gggg
            Input #1, hhhh
            ffff(dddd) = ffff(dddd) & vbCrLf & hhhh
        Next cccc
    Loop
    Close #1
    cccc = 0
    Do While Not (aaaa Or dddd < cccc)
        cccc = 1 - 0 + cccc
        Select Case LCase(bbbb)
        Case LCase(eeee(cccc))
            aaaa = True
        End Select
    Loop
    picResults.Cls
    If aaaa = True Then
        picResults.Print ffff(cccc)
    Else
        MsgBox "Please try again and make sure the term is spelt exactly as shown!"
    End If
End Sub

Private Sub cmdDefineCognitiveDisorders_Click()
    MsgBox "Cognitive Disorders are associated with an abnormal function in the brain in how information is acquired and processed as well as how it is stored."
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturntoDisorders_Click()
    frmCognitiveDisorders.Hide
    frmDisorders.Show
End Sub

Private Sub cmdReturnHome_Click()
    frmCognitiveDisorders.Hide
    frmHome.Show
End Sub


