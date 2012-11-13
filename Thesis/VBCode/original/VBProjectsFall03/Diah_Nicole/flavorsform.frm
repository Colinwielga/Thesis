VERSION 5.00
Begin VB.Form flavorsform 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Top flavors"
   ClientHeight    =   9660
   ClientLeft      =   2565
   ClientTop       =   780
   ClientWidth     =   10425
   LinkTopic       =   "Form2"
   ScaleHeight     =   9660
   ScaleWidth      =   10425
   Begin VB.CommandButton cmdquit2 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   27
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton cmdform1 
      Caption         =   "Go back to first form"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   26
      Top             =   7800
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   120
      Picture         =   "flavorsform.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1035
      TabIndex        =   25
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdflavorform 
      Caption         =   "Back to first page"
      Height          =   495
      Left            =   9960
      TabIndex        =   24
      Top             =   11400
      Width           =   1575
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7920
      TabIndex        =   23
      Top             =   11400
      Width           =   1575
   End
   Begin VB.CommandButton cmdinputfavorites 
      Caption         =   "Rate the flavors"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      TabIndex        =   20
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdreadfile 
      Caption         =   "Download"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      TabIndex        =   19
      Top             =   6360
      Width           =   1575
   End
   Begin VB.PictureBox picresults2 
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   1440
      ScaleHeight     =   3795
      ScaleWidth      =   5115
      TabIndex        =   18
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nicole Diah CS130"
      Height          =   375
      Left            =   8880
      TabIndex        =   29
      Top             =   9360
      Width           =   1575
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Rate your favorite ice cream flavors and then compare them to the official top 5 flavors at Ben and Jerry's"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   28
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Now enter your favorite flavors"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   22
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "First download the information"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   21
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFFF&
      Caption         =   "7) Vanilla for a Change:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFFF&
      Caption         =   "3)Chocolate Fudge Brownie:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFFF&
      Caption         =   "5) Chunky Monkey:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFFF&
      Caption         =   "9) Coffee for a Change"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFFF&
      Caption         =   "2)Chocolate Chip CookieDough:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFFF&
      Caption         =   "1) Cherry Garcia:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "8) Mint Chocolate Cookie:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "6) Butter Pecan:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "4) New York Super Fudge Chunk:"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000018&
      Caption         =   "Vanilla Ice Cream"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000018&
      Caption         =   "Chocolate Ice Cream with Fudge Brownies"
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000018&
      Caption         =   "Banana Ice Cream with Fudge Chunks and Walnuts "
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000018&
      Caption         =   "Coffee Ice Cream"
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000018&
      Caption         =   "Vanilla Ice Cream with Gobs of Chocolate Chip Cookie Dough"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000018&
      Caption         =   "Cherry Ice Cream with Cherries & Fudge Flakes"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000018&
      Caption         =   "Peppermint Ice Cream with Oxford Creme Cookies"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000018&
      Caption         =   "Rich Buttery Ice Cream with Roasted Pecans"
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "Chocolate Ice Cream with White and Dark Fudge Chunks, Pecans, Walnuts and Fudge Covered Almonds"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   2895
   End
End
Attribute VB_Name = "flavorsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name : benjerryproject.vbp (Nicole Diah's VB Project.vbp)
'Form Name : flavorsform (flavorsform.frm)
'Author: Nicole Diah
'Date Written: Oct. 27, 2003
'Purpose of Form: To display a list of the users top 5 favorite
                ' ice cream flavors and the official top 5 ice cream
                ' flavors at Ben and Jerry's

Dim icecreamflavor(1 To 13) As String, Path As String

Private Sub cmdform1_Click()
flavorsform.Hide
Form1.Show
End Sub

Private Sub cmdinputfavorites_Click()
Dim J As Integer, A As Integer, B As Integer, C As Integer, D As Integer, E As Integer
cmdreadfile.Enabled = False
picresults2.Cls   'clears the screen of any previous information
A = InputBox("Enter the number of the flavor you like the most", "Rate the flavors")
B = InputBox("Enter the number of your second favorite flavor?", "Rate the flavors")
C = InputBox("What is your third favorite flavor?", "Rate the flavors")
D = InputBox("What is your fourth favorite flavor?", "Rate the flavors")
E = InputBox("What is your fifth favorite flavor?", "Rate the flavors")

cmdinputfavorites.Enabled = False
picresults2.Print "Your favorite flavors" 'prints your list of favorites
picresults2.Print "****************************************************************************************************************************************"
picresults2.Print "1)", icecreamflavor(A)
picresults2.Print "2)", icecreamflavor(B)
picresults2.Print "3)", icecreamflavor(C)
picresults2.Print "4)", icecreamflavor(D)
picresults2.Print "5)", icecreamflavor(E)
picresults2.Print
picresults2.Print "Ben and Jerry's Top Ten flavors" ' a loop to print Ben and Jerry's official top 5 from the array
picresults2.Print "************************************************************************************************************************************************"
For J = 1 To 5
picresults2.Print J; ")", icecreamflavor(J)
Next J
cmdinputfavorites.Enabled = True

End Sub

Private Sub cmdquit2_Click()
End
End Sub

Private Sub cmdreadfile_Click()
Dim CTR As Integer
Open Path & "icecreamflavors.txt" For Input As #1 'open file
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, icecreamflavor(CTR)  'enters the information into an array
Loop
Close #1
cmdinputfavorites.Enabled = True

End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\Diah_Nicole\"
cmdinputfavorites.Enabled = False
cmdreadfile.Enabled = True
End Sub
