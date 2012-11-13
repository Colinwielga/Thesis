VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form3"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   11070
   LinkTopic       =   "Form3"
   ScaleHeight     =   8790
   ScaleWidth      =   11070
   Begin VB.CommandButton cmdDist 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to enter far away from the TwinCities you are willing to travel to college"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   5655
   End
   Begin VB.CommandButton cmdTuition 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to enter cost willing to spend on college tuition"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   5655
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go to the Next Slide"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back One Slide"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   3720
      Width           =   8175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Deciding on a College
' Form 2 (Tuition2)
' Kelsey Robinson
' March 10th, 2004
' This form prints the colleges that are less than what the user wants to pay on tuition
'and how far the user is willing to travel to the college.

Option Explicit


Private Sub cmdTuition_Click()
Found = False
Willing = InputBox("What are you willing to pay?")
picResults.Cls
picResults.Print "Name of the College"; Tab(35); "  Tuition" 'Tab(45); "Distance from the Twin Cities"
picResults.Print "************************************************************"
For J = 1 To CTR
    If Willing >= Tuition(J) Then
        picResults.Print College(J); Tab(35); FormatCurrency(Tuition(J))  'Tab(45); , Distance(J)
        Found = True
    End If
Next J
If Not Found Then
    picResults.Print "Sorry, there are no colleges with tuition as low as "; FormatCurrency(Willing)
End If
End Sub

Private Sub cmdDist_Click()
WillingDist = InputBox("How far (in miles) are you willing to travel?")
Found = False
picResults.Cls
picResults.Print "Name of the College"; Tab(30); "Distance from the Twin Cities in miles"
picResults.Print "***************************************************************************"
For J = 1 To CTR
    If WillingDist >= Distance(J) Then
        picResults.Print College(J); Tab(35); Distance(J)
        Found = True
    End If
Next J
If Not Found Then
    picResults.Print "Sorry, there are no colleges as close as "; WillingDist; " from the cities."
End If
End Sub
 
Private Sub cmdBack_Click()
Form3.Hide
Form2.Show
End Sub


Private Sub cmdNext_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

