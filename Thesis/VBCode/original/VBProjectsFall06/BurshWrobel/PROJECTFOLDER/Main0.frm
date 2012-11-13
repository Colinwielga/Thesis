VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8895
      Left            =   0
      Picture         =   "Main0.frx":0000
      ScaleHeight     =   8835
      ScaleWidth      =   10995
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.CommandButton Command6 
         Caption         =   "Quit"
         Height          =   495
         Left            =   5280
         TabIndex        =   12
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reviews/Favorites"
         Height          =   1215
         Left            =   240
         TabIndex        =   10
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Play Music!"
         Height          =   495
         Left            =   4920
         TabIndex        =   7
         Top             =   5280
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Credits/Order!"
         Height          =   1215
         Left            =   240
         TabIndex        =   6
         Top             =   6480
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Table of Contents"
         Height          =   1215
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enter Year"
         Height          =   1455
         Left            =   3960
         TabIndex        =   4
         Top             =   3600
         Width           =   3735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...When Finished"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "While Touring..."
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LETS BEGIN THE TOUR!"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   3240
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "*Dates are not precisely accurate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7920
         TabIndex        =   3
         Top             =   8520
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Megan Wrobel, Andrew Bursh - 11/02/06"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   8400
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "A History of Western Art"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   69
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   10935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form1
'Bursh, Wrobel
'11-01-06
'This form is the entry form into our program.  It provides the option for the user of how to search the works of art whether it be through the
'Enter Year button or the Table of Contents.  You can also goto a Review section (once your done viewing all) where you can choose you're favorites
'and it will list them in ascending order.  Dont forget to play music while you are touring!
    
    Option Explicit
Private Sub Command1_Click() 'Allows user to enter year from 25,000 BCE to 2006 AD, and will bring
'user to appropriate time period.
Dim Year As Integer

    Year = InputBox("Enter a year between [25,000 BCE - 2006 AD]...to discover artistic trends and styles from that year. Use (-) For BCE")
    Select Case Year
    Case -25000 To -6001
    Form2.Show
    Form1.Hide
    Case -6000 To -2500
    Form3.Show
    Form1.Hide
    Case -2499 To -451
    Form4.Show
    Form1.Hide
    Case -450 To -350
    Form5.Show
    Form1.Hide
    Case -349 To 0
    Form6.Show
    Form1.Hide
    Case 1 To 476
    Form7.Show
    Form1.Hide
    Case 477 To 850
    Form8.Show
    Form1.Hide
    Case 851 To 999
    Form9.Show
    Form1.Hide
    Case 1000 To 1149
    Form10.Show
    Form1.Hide
    Case 1150 To 1300
    Form11.Show
    Form1.Hide
    Case 1301 To 1450
    Form12.Show
    Form1.Hide
    Case 1451 To 1550
    Form13.Show
    Form1.Hide
    Case 1551 To 1600
    Form14.Show
    Form1.Hide
    Case 1600 To 1710
    Form15.Show
    Form1.Hide
    Case 1710 To 1750
    Form16.Show
    Form1.Hide
    Case 1751 To 1800
    Form17.Show
    Form1.Hide
    Case 1801 To 1850
    Form18.Show
    Form1.Hide
    Case 1851 To 1870
    Form19.Show
    Form1.Hide
    Case 1871 To 1899
    Form20.Show
    Form1.Hide
    Case 1900 To 1950
    Form21.Show
    Form1.Hide
    Case 1951 To 1960
    Form22.Show
    Form1.Hide
    Case 1961 To 2006
    Form23.Show
    Form1.Hide
    Case Else
    Print MsgBox("Date Outside of Range")
End Select

End Sub

Private Sub Command2_Click() 'Brings user to Table of Contents Form.
    frmTableOfContents.Show
    Form1.Hide
End Sub

Private Sub Command3_Click() 'Brings user to Order Form.
   Form25.Show
   Form1.Hide
End Sub

Private Sub Command4_Click() 'Brings user to Music Form
Print MsgBox("Are your speakers hooked up, with volume up?")
Form24.Show
Form1.Hide
End Sub

Private Sub Command5_Click() 'Brings user to Favorites Form.
Print MsgBox("Rate each Era 1-10 (1 = Terrible, 10 = Amazing), with whole numbers!")
Print MsgBox("You must enter a value for each Era!")
    Form26.Show
    Form1.Hide
End Sub

Private Sub Command6_Click()
End
End Sub
