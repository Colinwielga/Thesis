VERSION 5.00
Begin VB.Form formrates 
   BackColor       =   &H80000008&
   Caption         =   "Appealing Interest Rates"
   ClientHeight    =   8040
   ClientLeft      =   1320
   ClientTop       =   1950
   ClientWidth     =   12225
   LinkTopic       =   "Form2"
   ScaleHeight     =   8040
   ScaleWidth      =   12225
   Begin VB.CommandButton cmdhome 
      Caption         =   "Back To Home"
      Height          =   615
      Left            =   9480
      TabIndex        =   9
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdpersonal 
      Caption         =   "Personal Loans and Credit Lines"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   4440
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   3120
      Picture         =   "formrates.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   7560
      Picture         =   "formrates.frx":0BA4
      ScaleHeight     =   915
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin VB.PictureBox picwel 
      Height          =   975
      Left            =   6600
      Picture         =   "formrates.frx":2B07
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox picrates 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0FF&
      FillStyle       =   6  'Cross
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   1080
      ScaleHeight     =   1215
      ScaleWidth      =   10335
      TabIndex        =   2
      Top             =   6600
      Width           =   10335
   End
   Begin VB.CommandButton cmdnewcar 
      BackColor       =   &H00FF0000&
      Caption         =   "New Car Loans"
      DownPicture     =   "formrates.frx":354F
      BeginProperty Font 
         Name            =   "MS PMincho"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8520
      MaskColor       =   &H000000FF&
      Picture         =   "formrates.frx":433F
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   5760
      Picture         =   "formrates.frx":62A2
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"formrates.frx":73BA
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "    Wells Fargo Has Been      Found to Have the Best     Interest Rates Around!  "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1575
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "formrates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By :  Heather Parker
'form name = formrates
'form file name = Project One\formrates.frm
'Purpose : use Do While loops to pin point interest rates that
'appeal to what the user is looking for in a loan

Option Explicit
Dim months(1 To 5) As Integer
Dim Low(1 To 5) As Single
Dim high(1 To 5) As Single
Dim month As Integer
Dim Ctr As Integer
Dim J As Integer
Private Sub Form_Load()
Dim path As String
path = "N:\CS130\Parker, Heather\Project One\"
End Sub
'takes to form1
Private Sub cmdhome_Click()
formrates.Hide
Form1.Show
End Sub

Private Sub cmdnewcar_Click()
Dim path As String
Open path & "newcar.txt" For Input As #1
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, months(Ctr), Low(Ctr), high(Ctr)
Loop
Close #1
month = InputBox("How many Months Would you Like Your Loan to Extend?")
J = 1
'looks for the number of months with in a range that the user would like their loan
'for and matches to that to an interest rate and amount of time
Do While month > months(Ctr)
    J = J + 1
Loop
picrates.Print "A loan extending"; month; "months will have rates ranging from"; Low(Ctr); "%"; "to"; high(Ctr); "% for a New Car"
End Sub

Private Sub cmdpersonal_Click()
Dim amount(1 To 4) As Single
Dim years(1 To 4) As Integer
Dim rate(1 To 4) As Single
Dim time As Integer
Dim J As Integer
Dim amounted As Single
Dim path As String
Open path & "personalloans.txt" For Input As #1
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, amount(Ctr), years(Ctr), rate(Ctr)
Loop
Close #1
amounted = InputBox("How much would you like to Loan Out (Loan rates are from 5-10 Years)?")
Ctr = 1
'looks for an amount from the user and matches that within a range to come up with an interest rate for the loan
Do While amounted > amount(Ctr)
    Ctr = Ctr + 1
Loop
picrates.Print "A personal loan for the amount of"; FormatCurrency(amounted); "will have an interest rate of"; rate(Ctr); "%"

End Sub


