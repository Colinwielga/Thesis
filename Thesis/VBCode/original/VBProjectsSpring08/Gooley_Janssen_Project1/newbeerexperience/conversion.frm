VERSION 5.00
Begin VB.Form frmConversions 
   BackColor       =   &H00800000&
   Caption         =   "Conversions"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0000FFFF&
      Caption         =   "CLICK FIRST to Read File"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   3360
      ScaleHeight     =   915
      ScaleWidth      =   7755
      TabIndex        =   11
      Top             =   6480
      Width           =   7815
   End
   Begin VB.TextBox txtDesired 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   3480
      TabIndex        =   10
      Top             =   5040
      Width           =   7335
   End
   Begin VB.TextBox txtOriginal 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   3480
      TabIndex        =   9
      Top             =   3720
      Width           =   7335
   End
   Begin VB.TextBox txtAmount 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   3480
      TabIndex        =   8
      Top             =   2280
      Width           =   7215
   End
   Begin VB.Label lblConverted 
      BackColor       =   &H00800000&
      Caption         =   "Converted Amount is: "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   480
      TabIndex        =   7
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label lblDesired 
      BackColor       =   &H00800000&
      Caption         =   "Desired Unit (1 - 4)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   480
      TabIndex        =   6
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label lblOriginal 
      BackColor       =   &H00800000&
      Caption         =   "Original Unit (1 - 4)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00800000&
      Caption         =   "Amount to Convert"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lblKeg 
      BackColor       =   &H00800000&
      Caption         =   "4. Keg"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   11880
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblLiter 
      BackColor       =   &H00800000&
      Caption         =   "3. Liter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   9840
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblPint 
      BackColor       =   &H00800000&
      Caption         =   "2. Pint"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   8040
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblOunce 
      BackColor       =   &H00800000&
      Caption         =   "1. Ounce"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmConversions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Beer Experience
'frm Conversions
'Lauren Gooley and Tim Janssen
'March 22, 2008
'This form allows the user to convert from one unit of liquid to another desired unit of liquid (i.e. ounces to pints).

Option Explicit

Private Sub cmdMainMenu_Click()
frmConversion.Hide
frmStartUp.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()
Dim CTR As Integer, A As Single, O As Single, D As Single, Value As Single, conversions(1 To 100) As Single
Open App.Path & "\conversions.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, conversions(CTR)
Loop
A = txtAmount.Text
O = txtOriginal.Text
D = txtDesired.Text
Value = ((A * conversions(O)) / conversions(D))
picResults.Print FormatNumber(Value, 3)
End Sub
End Sub
