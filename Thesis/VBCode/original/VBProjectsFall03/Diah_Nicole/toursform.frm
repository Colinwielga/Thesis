VERSION 5.00
Begin VB.Form toursform 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Factory Tour Information"
   ClientHeight    =   8475
   ClientLeft      =   3105
   ClientTop       =   1650
   ClientWidth     =   9495
   LinkTopic       =   "Form2"
   ScaleHeight     =   8475
   ScaleWidth      =   9495
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   240
      Picture         =   "toursform.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   21
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdfirstform 
      Caption         =   "Back to first form"
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
      Left            =   5040
      TabIndex        =   20
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdquit 
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
      Left            =   6960
      TabIndex        =   19
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find  Information     Now"
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
      Left            =   5760
      TabIndex        =   18
      Top             =   5400
      Width           =   2415
   End
   Begin VB.PictureBox Picture4 
      Height          =   855
      Left            =   5760
      Picture         =   "toursform.frx":0D78
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   17
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   480
      Picture         =   "toursform.frx":166D
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   16
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   3000
      Picture         =   "toursform.frx":1F62
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   15
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox txtchildren 
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox txtseniors 
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox txtadults 
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox txtmonth 
      Height          =   405
      Left            =   3840
      TabIndex        =   4
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox txtday 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox picresults4 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   360
      ScaleHeight     =   2955
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nicole Diah  CS130"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CHILDREN    (12 or younger)"
      Height          =   495
      Left            =   7320
      TabIndex        =   14
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SENIORS     (65 or older)"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADULTS"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter the number of visitors:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "enter (day / month)"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Date for the tour:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Factory Tour Information:"
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
      TabIndex        =   2
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"toursform.frx":2857
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "toursform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name : benjerryproject.vbp (Nicole Diah's VB Project.vbp)
'Form Name : toursform (toursform.frm)
'Author: Nicole Diah
'Date Written: Oct. 27, 2003
'Purpose of Form: To display the tour hours available for a certain date
                ' given by the user. Also to calculate the admission costs
                ' for a given amount of guests
Dim Path As String

Private Sub cmdfind_Click()
Dim month As Integer, CTR As Integer, adult As Single, senior As Single, child As Single, tours(1 To 6) As String, tourtime(1 To 6) As String
Dim shophours(1 To 6) As String, adultprice As Single, seniorprice As Single, childprice As String, total As Single, day As Integer
picresults4.Cls 'clears the screen of previous information
CTR = 0
month = txtmonth.Text
day = txtday.Text
Open Path & "tourschedule.txt" For Input As #1 'used to enter the monthly tour information in 3 arrays
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tours(CTR), tourtime(CTR), shophours(CTR)
Loop
Close #1
picresults4.Print "Information for date:"; day; "/"; month
picresults4.Print
picresults4.Print tours(1), tourtime(1), shophours(1)
picresults4.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Select Case month
    Case 1, 2, 3, 4, 5, 11, 12
        picresults4.Print tours(2), tourtime(2); Tab(48); shophours(2) 'prints the information for the date entered by user
    Case 6
        picresults4.Print tours(3), , tourtime(3); Tab(48); shophours(3)
    Case 7, 8
        picresults4.Print tours(4), , tourtime(4); Tab(48); shophours(4)
    Case 9, 10
        picresults4.Print tours(5), , tourtime(5); Tab(48); shophours(5)
    Case Is > 12
        picresults4.Print "unacceptable date entered"
End Select

picresults4.Print "*************************************************************************************************************************************************************************"
adult = txtadults.Text
senior = txtseniors.Text
child = txtchildren.Text 'gets information from the user from the textboxes
adultprice = 3 * adult
seniorprice = 2 * senior
childprice = "FREE"
total = adultprice + seniorprice
picresults4.Print "ADMISSION COSTS:"
picresults4.Print "adults:", FormatCurrency(adultprice) 'print the prices for each age group
picresults4.Print "seniors:", FormatCurrency(seniorprice)
picresults4.Print "children:", childprice
picresults4.Print "__________________"
picresults4.Print "Total:", FormatCurrency(total) 'print the total
    
End Sub

Private Sub cmdfirstform_Click()
Form1.Show
toursform.Hide
nutritionform.Hide
flavorsform.Hide
'hides all forms except first form
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\Diah_Nicole\"
cmdfind.Enabled = False
End Sub

Private Sub txtchildren_Change()
cmdfind.Enabled = True
End Sub
