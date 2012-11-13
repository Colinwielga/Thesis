VERSION 5.00
Begin VB.Form frmTedStore 
   Caption         =   "Ted's Store"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   Picture         =   "frmTedStore.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbegin 
      BackColor       =   &H0000FFFF&
      Caption         =   "CLICK HERE TO BEGIN PURCHASE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9960
      Width           =   4095
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8640
      Width           =   4095
   End
   Begin VB.CommandButton cmdbiobook 
      Caption         =   " "
      Height          =   3015
      Left            =   7200
      Picture         =   "frmTedStore.frx":2EBED
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   3135
   End
   Begin VB.CommandButton cmdbobble 
      Height          =   3015
      Left            =   4920
      Picture         =   "frmTedStore.frx":33F68
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton cmdtee 
      Caption         =   " "
      DownPicture     =   "frmTedStore.frx":475CE
      Height          =   3255
      Left            =   7200
      Picture         =   "frmTedStore.frx":504E1
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   3135
   End
   Begin VB.CommandButton cmdslugger 
      Caption         =   " "
      Height          =   3015
      Left            =   240
      Picture         =   "frmTedStore.frx":593F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdmonoploy 
      Caption         =   " "
      Height          =   3015
      Left            =   2640
      Picture         =   "frmTedStore.frx":693FB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton cmdplate 
      Caption         =   " "
      Height          =   3255
      Left            =   4920
      Picture         =   "frmTedStore.frx":725EF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      Height          =   5775
      Left            =   10920
      ScaleHeight     =   5715
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   2760
      Width           =   4095
   End
   Begin VB.CommandButton cmdbatbook 
      Caption         =   " "
      Height          =   3255
      Left            =   2640
      Picture         =   "frmTedStore.frx":82111
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Cmdreplica 
      Height          =   3255
      Left            =   240
      Picture         =   "frmTedStore.frx":8D281
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdTedStoreBack 
      BackColor       =   &H80000009&
      Caption         =   "Click to go back to Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9480
      Width           =   4695
   End
End
Attribute VB_Name = "frmTedStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim Sum As Single


Private Sub cmdbatbook_Click()
    Dim cmdbatbook As Single
    cmdbatbook = 24.95
    picresults.Print "Batting The Ted Williams Way"; Tab; FormatCurrency(cmdbatbook)
    Sum = Sum + cmdbatbook
End Sub

Private Sub cmdbegin_Click()
    A = InputBox("Enter Your Name for Our Records", "The Ted Willams Store")
    MsgBox "CLICK ON ITEMS YOU WOULD LIKE TO PURCHASED", , "NOTICE"
End Sub

Private Sub cmdbiobook_Click()
    Dim cmdbiobook As Single
    cmdbiobook = 24.95
    picresults.Print "A Portrait in Words and Pictures"; Tab; FormatCurrency(cmdbiobook)
    Sum = Sum + cmdbiobook
End Sub

Private Sub cmdbobble_Click()
    Dim cmdbobble As Single
    cmdbobble = 35
    picresults.Print "Ted Williams Bobble Heads"; Tab; Tab; FormatCurrency(cmdbobble)
    Sum = Sum + cmdbobble
End Sub

Private Sub cmdmonoploy_Click()
    Dim cmdmonoploy As Single
    cmdmonoploy = 35
    picresults.Print "Boston Red Sox Monopoly"; Tab; Tab; FormatCurrency(cmdmonoploy)
    Sum = Sum + cmdmonoploy
End Sub

Private Sub cmdplate_Click()
    Dim cmdplate As Single
    cmdplate = 225
    picresults.Print " Bradford Ex - Ceramic Plate"; Tab; Tab; FormatCurrency(cmdplate)
    Sum = Sum + cmdplate
End Sub

Private Sub Cmdreplica_Click()
    Dim Cmdreplica As Single
    Cmdreplica = 220
    picresults.Print "Red Sox Replica Jersey"; Tab; Tab; FormatCurrency(Cmdreplica)
    Sum = Sum + Cmdreplica
End Sub

Private Sub cmdslugger_Click()
    Dim cmdslugger As Single
    cmdslugger = 3995
    picresults.Print "Louisville Slugger - Black Bat"; Tab; FormatCurrency(cmdslugger)
    Sum = Sum + cmdslugger
End Sub

Private Sub cmdTedStoreBack_Click()
    frmTedmenu.Show
    frmTedStore.Hide
End Sub

Private Sub cmdtee_Click()
    Dim cmdtee As Single
    cmdtee = 14.95
    picresults.Print "Ted Williams signature t-shirt"; Tab; Tab; FormatCurrency(cmdtee)
    Sum = Sum + cmdtee
End Sub

Private Sub cmdtotal_Click()
Dim total As Integer, tax As Integer, shipping As Integer
    tax = Sum * (6 / 100)
    shipping = 7
    total = Sum + shipping + tax
    picresults.Print ""
    picresults.Print ""
    picresults.Print A; "'s Purchase today"
    picresults.Print "*********************************************************************"
    picresults.Print "Sub-Total"; Tab; Tab; Tab; FormatCurrency(Sum)
    picresults.Print "Tax"; Tab; Tab; Tab; FormatCurrency(tax)
    picresults.Print "Shipping Cost"; Tab; Tab; Tab; FormatCurrency(shipping)
    picresults.Print "*********************************************************************"
    picresults.Print "TOTAL"; Tab; Tab; Tab; FormatCurrency(total)
    picresults.Print "*********************************************************************"
    picresults.Print "*********************************************************************"
End Sub

Private Sub Command9_Click()
    picresults.Cls
    Sum = 0
    total = 0
End Sub
