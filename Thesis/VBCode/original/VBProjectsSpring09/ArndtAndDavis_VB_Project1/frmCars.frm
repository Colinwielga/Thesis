VERSION 5.00
Begin VB.Form frmCars 
   BackColor       =   &H00000000&
   Caption         =   "Sweet Deals On Wheels"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGuessAnswers 
      BackColor       =   &H000000FF&
      Caption         =   "Check My Guesses"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   30
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   -120
      Picture         =   "frmCars.frx":0000
      ScaleHeight     =   2055
      ScaleWidth      =   3615
      TabIndex        =   29
      Top             =   2760
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   9960
      ScaleHeight     =   3075
      ScaleWidth      =   4635
      TabIndex        =   27
      Top             =   1800
      Width           =   4695
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit Vehicle Choice and Continue to Test Your Luck"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9960
      TabIndex        =   26
      Top             =   7200
      Width           =   4935
   End
   Begin VB.TextBox txtCar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10320
      TabIndex        =   25
      Top             =   6240
      Width           =   3855
   End
   Begin VB.TextBox txtFour 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5520
      TabIndex        =   18
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txtThree 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5520
      TabIndex        =   17
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtFive 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5520
      TabIndex        =   16
      Top             =   9720
      Width           =   1215
   End
   Begin VB.TextBox txtOne 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5520
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtTwo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5520
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
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
      Left            =   10200
      TabIndex        =   7
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton cmdHomepage 
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
      Left            =   11520
      TabIndex        =   6
      Top             =   9720
      Width           =   1095
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
      TabIndex        =   5
      Top             =   9720
      Width           =   1095
   End
   Begin VB.PictureBox picE 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -120
      Picture         =   "frmCars.frx":3BCA
      ScaleHeight     =   2295
      ScaleWidth      =   3615
      TabIndex        =   3
      Top             =   6840
      Width           =   3615
   End
   Begin VB.PictureBox picD 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   -120
      Picture         =   "frmCars.frx":5B06
      ScaleHeight     =   2175
      ScaleWidth      =   3615
      TabIndex        =   2
      Top             =   4680
      Width           =   3615
   End
   Begin VB.PictureBox picC 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   -120
      Picture         =   "frmCars.frx":77F0
      ScaleHeight     =   2535
      ScaleWidth      =   3615
      TabIndex        =   1
      Top             =   8400
      Width           =   3615
   End
   Begin VB.PictureBox picB 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      Picture         =   "frmCars.frx":908B
      ScaleHeight     =   1815
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label lblResults 
      BackColor       =   &H00000000&
      Caption         =   "YOUR RESULTS"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   11160
      TabIndex        =   28
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Enter the Name of Your Dream Vehicle:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   10800
      TabIndex        =   24
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label lbl340000 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E. $340,000"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7200
      TabIndex        =   23
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label lbl6995 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D. $6,995"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7200
      TabIndex        =   22
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lbl21790 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C. $21,790"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7200
      TabIndex        =   21
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label lbl6000 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B. $6,000"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7200
      TabIndex        =   20
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lbl150 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A. $150"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7200
      TabIndex        =   19
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblH 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Enter the letter of the price (A-E) from below into the text box by each vehicle to guess the cost!"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   480
      Width           =   14055
   End
   Begin VB.Label lblG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "4: Harvey Davidson Motorcycle"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   3480
      TabIndex        =   12
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label lblF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "3: Chevy Cavalier"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3960
      TabIndex        =   11
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblE 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "5: VW Jetta"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4080
      TabIndex        =   10
      Top             =   9840
      Width           =   1455
   End
   Begin VB.Label lblD 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "1: Rolls- Royce"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3840
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "2: Bike"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblA 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Match the vehicle with the appropriate price!"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "frmCars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmCars
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective: On this form, User plays matching game for prices with cars with an If-Then Conditional statement, then chooses a car which is saved for their summary

Option Explicit

Private Sub cmdBack_Click()
frmTheFinerThingsInLife.Show
frmCars.Hide
End Sub

Private Sub cmdGuessAnswers_Click()
'This button will check the users input guesses to match price and cars

picResults.Cls
If ((txtOne.Text = "E" Or txtOne.Text = "e") And (txtTwo.Text = "A" Or txtTwo.Text = "a") And (txtThree.Text = "D" Or txtThree.Text = "d") And (txtFour.Text = "B" Or txtFour.Text = "b") And (txtFive.Text = "C" Or txtFive.Text = "c")) Then
    picResults.Print "CORRECT! Great Job!"
    picResults.Print
        picResults.Print
    picResults.Print "Type your dream vehicle "
    picResults.Print "into the textbox below"
    End If
If txtOne.Text <> "E" And txtOne.Text <> "e" Then
    picResults.Print "Incorrect guess for Rolls-Royce"
    End If
If txtTwo.Text <> "A" And txtTwo.Text <> "a" Then
    picResults.Print "Incorrect guess for Bike"
    End If
If txtThree.Text <> "D" And txtThree.Text <> "d" Then
    picResults.Print "Incorrect guess for Chevy Cavalier"
    End If
If txtFour.Text <> "B" And txtFour.Text <> "b" Then
    picResults.Print "Incorrect guess for Motorcycle"
    End If
If txtFive.Text <> "C" And txtFive.Text <> "c" Then
    picResults.Print "Incorrect guess for VW Jetta"
    End If
    
    
End Sub

Private Sub cmdHomepage_Click()
frmBeginning.Show
frmCars.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSubmit_Click()
Car = txtCar.Text
frmLucky.Show
frmCars.Hide
End Sub

