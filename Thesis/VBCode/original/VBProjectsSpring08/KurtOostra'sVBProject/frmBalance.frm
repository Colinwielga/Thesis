VERSION 5.00
Begin VB.Form frmBalance 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3375
      Left            =   1080
      ScaleHeight     =   3315
      ScaleWidth      =   8835
      TabIndex        =   13
      Top             =   3960
      Width           =   8895
   End
   Begin VB.CommandButton cmdBricks 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate the most effecient use of bricks to balance the system."
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7680
      Width           =   2895
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate the weight to be added."
      Height          =   1095
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   2535
   End
   Begin VB.TextBox txtCyc 
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtS4Par 
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtPars 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtFresnel 
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtSource4 
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      Height          =   1095
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Fill each box, add a zero if no lights are being added"
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Enter the number of Cyc lights used:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "Enter the number of Source 4 Par lights used:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Enter the number of Par Can lights used:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Enter the number of Fresnel lights used:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the number of Source 4 lights used:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   4935
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Theater Lighting
'Form Name: frmBalance
'Author: Kurt Oostra
'Date Written: 3/12/08
'Objective: Calculate the weight needed to balance a lineset.
'           Determine the best usage of bricks to balance the system.
Option Explicit
Dim Sum As Single
Dim S4 As Single, S4Par As Single, Cyc As Single, Fres As Single, Par As Single

Private Sub cmdBricks_Click()
Dim z As Integer, i As Integer, x As Integer
x = 0
i = 0
z = 0
'Determine the best usage of bricks based on the balance weight minus the bricks.
'Bricks come in sizes of 35 lbs, 25 lbs, and 12.5 lbs
Do While Sum >= 35
    Sum = Sum - 35
    i = i + 1
Loop
Do While Sum >= 25
    Sum = Sum - 25
    x = x + 1
Loop
Do While Sum >= 12.5
    Sum = Sum - 12.5
    z = z + 1
Loop
'Prints only the number of bricks used when greater then 0
If i > 0 Then picResults.Print i; " 35 lb bricks should be used"
If x > 0 Then picResults.Print x; " 25 lb bricks should be used"
If z > 0 Then picResults.Print z; " half bricks should be used."
End Sub

Private Sub cmdCalc_Click()
S4Par = txtS4Par
Fres = txtFresnel
S4 = txtSource4
Par = txtPars
Cyc = txtCyc
'clears pic box
picResults.Cls
'determines weight needed to balance the lineset
Sum = S4 * 20 + S4Par * 12.8 + Fres * 21 + Cyc * 30 + Par * 11
'prints this weight
picResults.Print "Approximately "; Sum; " lbs should be added to balance the system."
'enables the brick button
cmdBricks.Enabled = True
End Sub

Private Sub cmdReturn_Click()
'returns to main menu
frmMainMenu.Show
frmBalance.Hide
End Sub
