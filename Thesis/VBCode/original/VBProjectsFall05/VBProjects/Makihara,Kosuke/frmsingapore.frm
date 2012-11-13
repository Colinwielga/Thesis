VERSION 5.00
Begin VB.Form frmsingapore 
   Caption         =   "Singapore"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdload2 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdrank 
      Caption         =   "See the Top 10"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Display"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   17
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdload 
      Cancel          =   -1  'True
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdidol 
      Caption         =   "See the Selections"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdquitsin 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get  Result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtlang 
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdethnic 
      Caption         =   "Get Resut"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtethnic 
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton cmdanswersin 
      Caption         =   "Get Resut"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdgdpsin 
      Caption         =   "What is GDP??"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtgdpsin 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.PictureBox picoutputsin 
      Height          =   2655
      Left            =   4920
      ScaleHeight     =   2595
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   3960
      Width           =   4455
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label labidol 
      BackColor       =   &H80000002&
      Caption         =   "Which is the winner of ""Singapore Idol""??"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label lbllang 
      BackColor       =   &H80000002&
      Caption         =   "Which languages are spoken in Singapore??      Hint) see the answer of the culture "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label lblimg 
      BackColor       =   &H80000002&
      Caption         =   "Singapre is also such as multi-cultural place. Which culture are there in Singapore?             Hint) almost overall Asia...    "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label lblgdpsin 
      BackColor       =   &H80000002&
      Caption         =   "How much is the recent GDP of Singapore?? Hint) US: 11trillion 750billion"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Singapore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image Image4 
      Height          =   7200
      Left            =   0
      Picture         =   "frmsingapore.frx":0000
      Top             =   0
      Width           =   9600
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   5160
      Top             =   600
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   5160
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmsingapore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Tokyo, Berlin, Singapore- My Summer 2005 (Makihara_Kosuke.vbp)
'Form Name: Singapore(frmsingapore.frm)
'Author: Kosuke Makihara
'Date Wrriten: 27 Oct 2005
'Ojectives:
'This form introduce Singapore and provide some trivia questions
'about the city.
Option Explicit
Dim nation(1 To 11) As String
Dim X As Single
Dim gdp(1 To 11) As Single
Dim country(1 To 6) As String
Dim number(1 To 6) As Single
Private Sub cmdquit_Click()
End
End Sub
Private Sub Image1_Click()
End Sub
Private Sub cmdanswersin_Click()
'This button determine how user's estimate of singapore's
'GDP is close to the each country's GDP.

Dim answer As Single
answer = txtgdpsin.Text
X = 1
Do Until answer > gdp(X)
X = X + 1
Loop
picoutputsin.Print gdp(X), nation(X)



End Sub

Private Sub cmdclear_Click()
'This button clear the output in the picture box.
picoutputsin.Cls

End Sub

Private Sub cmdethnic_Click()
'This button works to determine how user's estimate about the
'ethnicity in sigapore, and show how many each groups there are
'in the city.

Dim national As String
national = txtethnic.Text
X = 1
Do Until national = country(X)
    X = X + 1
    Loop
 picoutputsin.Print country(X), FormatPercent(number(X))
 
   
    
    
End Sub

Private Sub cmdgdpsin_Click()
'Explain about GDP by message box.
MsgBox "GDP: Gross Domestic Production...GDP is the total value of goods and services produced by a nation.", , "GDP"
End Sub

Private Sub cmdidol_Click()
'This button take users to the new form.
frmsinidol.Show
End Sub

Private Sub cmdload_Click()
'This button is for loading the data about ethnicity in Singapore.
'from the file.
Open App.Path & "\sin_ethnic.txt" For Input As #1
For X = 1 To 6
    Input #1, country(X), number(X)
Next X
End Sub

Private Sub cmdload2_Click()
'This button is for loadning the GDP data from the file.
Open App.Path & "\gdp.txt" For Input As #1
For X = 1 To 11
    Input #1, gdp(X), nation(X)
Next X


End Sub

Private Sub cmdmain_Click()
'This button take user to the main form.
frmsingapore.Hide
frmmain.Show
End Sub

Private Sub cmdquitsin_Click()
End
End Sub

Private Sub cmdrank_Click()
'This conde shows the top ten of highest GDP in the world.
For X = 1 To 11
picoutputsin.Print gdp(X), nation(X)
Next X



End Sub

Private Sub Command1_Click()
'This code determine the user's input about the spoken language
'singapore. Depending on the popularity of language, the message
'differs.


Dim lang As String

lang = txtlang.Text
Select Case lang

Case Is = "English"
picoutputsin.Print "Yes, English is widely spoken in Singapore"

Case Is = "Chinese"
picoutputsin.Print "Yes, Chinese is spoken in Singapore"

Case Is = "Malay", "Tamir", "Hindu", "Arabic"
picoutputsin.Print "Yes, this language is spoken only among certain group of people."
Case Else
picoutputsin.Print "Not really...thinks of other languages"
End Select

End Sub




