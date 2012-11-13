VERSION 5.00
Begin VB.Form frmInputData 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Find Your Position(s)"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to the Home Page"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   3735
   End
   Begin VB.CommandButton cmdFindPos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find Your Position!"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      Picture         =   "frmInputData.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   5895
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6600
      TabIndex        =   4
      Text            =   "0"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9480
      TabIndex        =   3
      Text            =   "0"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtForty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      TabIndex        =   2
      Text            =   "0"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtAccScore 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9480
      TabIndex        =   1
      Text            =   "0"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Throwing Accuracy Score"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   8880
      TabIndex        =   9
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Forty Yard Dash Time"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   6240
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Weight"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Height in inches"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   6240
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "frmInputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   Input Data
'   Greg Sowden
'   10/9/09
'   This form allows the user to input his/her own data into the machine to get an output of what position he/she should play
'   the form uses text boxes for input boxes

'   first, show the input data form and hide the home
Private Sub cmdBack10_Click()
    frmInputData.Hide
    frmRoster.Show
    
End Sub

Private Sub cmdFindPos_Click()
    Dim forty As Single, tall As Single, heavy As Integer, name As String, acc As Integer
'   set variables equal to text names to use them in formula
    name = txtName.Text
    forty = txtForty.Text
    acc = txtAccScore.Text
    tall = txtHeight.Text
    heavy = txtWeight.Text
'   set if statements to give msgbox outputs of the positions.  By doing seperate if statements, the user gets multiple msgboxes, one for every position
    If forty >= 5.3 And heavy >= 240 Then
        MsgBox name & " with a weight of " & heavy & " pounds and a forty time of " & forty & " seconds would be best suited as an Offensive Linemen", , "Your Position!"
    End If
    
    If forty <= 5.3 And heavy > 210 Then
        MsgBox name & " with a weight of " & heavy & " pounds, a height of " & tall & " inches, and a forty time of " & forty & " seconds would be best suited as a Fullback", , "Your Position!"
    End If
    
    If forty < 5 And heavy <= 240 Then
        MsgBox name & " with a weight of " & heavy & " pounds and a forty time of " & forty & " seconds would be best suited as a Runningback", , "Your Position!"
    End If
    
    If forty <= 5.3 And heavy <= 230 And acc > 70 And tall > 68 Then
         MsgBox name & " with a weight of " & heavy & " pounds, a height of " & tall & " inches, a throwing accuracy score of " & acc & " and a forty time of " & forty & " seconds would be best suited as a Quarterback", , "Your Position!"
    End If
    
    If forty <= 5 And heavy <= 230 And tall >= 72 Then
         MsgBox name & " with a weight of " & heavy & " pounds, a height of " & tall & " inches, and a forty time of " & forty & " seconds would be best suited as a Wide Reciever", , "Your Position!"
    End If
    
    If forty >= 5.3 And heavy < 240 Then
        MsgBox "The statistics do not line up well with any position", , "No Position Found"
    
    End If
        
End Sub

