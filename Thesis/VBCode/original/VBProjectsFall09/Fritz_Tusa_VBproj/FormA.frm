VERSION 5.00
Begin VB.Form Airfare 
   Caption         =   "Airfare"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13920
   LinkTopic       =   "Drive"
   ScaleHeight     =   10065
   ScaleWidth      =   13920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResultsAirfare 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7200
      ScaleHeight     =   1755
      ScaleWidth      =   4995
      TabIndex        =   10
      Top             =   6720
      Width           =   5055
   End
   Begin VB.CommandButton cmdAirfare 
      BackColor       =   &H000080FF&
      Caption         =   "Calculate!!"
      Height          =   1095
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox firstclasstext 
      Height          =   735
      Left            =   9720
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox many 
      Height          =   855
      Left            =   9720
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox state 
      Height          =   855
      Left            =   9720
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Century"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Text            =   "Calculate the Cost of Airfare to Colorado!"
      Top             =   360
      Width           =   8655
   End
   Begin VB.CommandButton cmdToTitle 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Title"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Will you be flying first class? (1=Yes/0=No)"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Please enter how many people will be in your group =====>"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   6
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Enter the Color of your state here, for international fliers, please put Orange. ===============>"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"FormA.frx":0000
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   4395
      Left            =   0
      Picture         =   "FormA.frx":046D
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   6045
   End
   Begin VB.Image Image2 
      Height          =   10290
      Left            =   0
      Picture         =   "FormA.frx":549E3
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   13920
   End
End
Attribute VB_Name = "Airfare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SKI TRIP'
'AIRFARE'
'MAX TUSA'
'8-18'
'THIS FORM ALLOWS THE USER TO CALCULATE THE COST OF AIRFARE FROM A SPECIFIC LOCATION'

Option Explicit

Private Sub cmdAirfare_Click()
'this finds the cost of airfare for an idividual'
'dim variables'
Dim numberoftickets As Integer, firstclass As Integer, ColorState As String
Dim cost As Currency, wherefrom As String, allairfare As Currency

'set airfare cost to zero to guard against multiple entries'
totalairfarecost = 0

'clear any leftover images in the pictuer box'
picResultsAirfare.Cls

'warning message box'
MsgBox "Make sure you enter the correct color that corresponds to where you are flying from, *REMEMBER, COLORS ARE CASE SENSITIVE!*", , "DONT FORGET"

'indicate where input is to come from'
ColorState = state.Text
firstclass = firstclasstext.Text
numberoftickets = many.Text

'select case for the area you are flying from'
Select Case ColorState
    Case "Green"
        cost = 200
        wherefrom = " The Central U.S.A "
    Case "Red"
        cost = 200
        wherefrom = " The West Coast "
    Case "Orange"
        cost = 200
        wherefrom = " A long ways away "
    Case "Blue"
        cost = 100
        wherefrom = " Close_enough_to_drive "
    Case "Yellow"
        cost = 300
        wherefrom = " The East Coast "
    Case "Purple"
        cost = 300
        wherefrom = " Alaska! "
    Case Else
        MsgBox "TRY AGAIN", , "FAIL"
End Select

'Create and if then statement for if first class is selectd'
If firstclass = 1 Then
    cost = cost + (cost * 0.5)
ElseIf firstclass = 0 Then
    cost = cost
ElseIf firstclass < 0 And firstclass > 1 Then
    MsgBox "TRY AGAIN", , "FAIL"
End If
    
'make a header'
picResultsAirfare.Print "COST OF AIRFARE FROM"; Tab(1); wherefrom; Tab(1); "WITH"; numberoftickets; "PEOPLE"

'calculate the cost'
allairfare = cost * numberoftickets

'add to total cost'
totalairfarecost = allairfare

'display the cost'
picResultsAirfare.Print ""
picResultsAirfare.Print FormatCurrency(allairfare)

End Sub

Private Sub cmdToTitle_Click()
Title.Show
Airfare.Hide
End Sub

