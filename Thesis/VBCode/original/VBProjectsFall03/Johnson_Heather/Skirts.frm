VERSION 5.00
Begin VB.Form frmSkirts1 
   BackColor       =   &H008080FF&
   Caption         =   "Skirts"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   13290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CLEAR"
      Height          =   855
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "QUIT!!!!"
      Height          =   855
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdmainmenu 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Back to the Main Menu!!!"
      Height          =   1215
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmddiscount 
      BackColor       =   &H00C0C0FF&
      Caption         =   "What is the total after the discount??"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.PictureBox picSkirttotal 
      Height          =   3495
      Left            =   6360
      ScaleHeight     =   3435
      ScaleWidth      =   4035
      TabIndex        =   21
      Top             =   5640
      Width           =   4095
   End
   Begin VB.CommandButton cmdskirttotal 
      BackColor       =   &H00C0C0FF&
      Caption         =   "What's the total for the skits?"
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox txthowmanyskirts 
      Height          =   615
      Left            =   6600
      TabIndex        =   19
      Top             =   4800
      Width           =   3375
   End
   Begin VB.OptionButton optncc 
      BackColor       =   &H008080FF&
      Height          =   375
      Left            =   11400
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optflyers 
      BackColor       =   &H008080FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optflyaway 
      BackColor       =   &H008080FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optaline 
      BackColor       =   &H008080FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picNCC 
      Height          =   1695
      Left            =   10440
      Picture         =   "Skirts.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.PictureBox picFlyers 
      Height          =   1695
      Left            =   7800
      Picture         =   "Skirts.frx":4A9F
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.PictureBox picFlyaway 
      Height          =   1815
      Left            =   5280
      Picture         =   "Skirts.frx":8F07
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.PictureBox picAline 
      Height          =   1695
      Left            =   2640
      Picture         =   "Skirts.frx":9820
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblstep5 
      BackColor       =   &H008080FF&
      Caption         =   "STEP 5: BACK TO THE MAIN MENU"
      BeginProperty Font 
         Name            =   "SchoolBoy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11280
      TabIndex        =   30
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label lblstep4 
      BackColor       =   &H008080FF&
      Caption         =   "STEP 4: Find out what your total is after your discount"
      BeginProperty Font 
         Name            =   "SchoolBoy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   29
      Top             =   8040
      Width           =   3495
   End
   Begin VB.Label lblstep3 
      BackColor       =   &H008080FF&
      Caption         =   "STEP 3: Find out what your total is before the discount"
      BeginProperty Font 
         Name            =   "SchoolBoy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   28
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Label lblstep2 
      BackColor       =   &H008080FF&
      Caption         =   "STEP 2: Enter how many skirts you will need"
      BeginProperty Font 
         Name            =   "SchoolBoy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   27
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label lblstep1 
      BackColor       =   &H008080FF&
      Caption         =   "STEP 1:  CHOOSE A SKIRT STYLE"
      BeginProperty Font 
         Name            =   "SchoolBoy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblhowmany 
      BackColor       =   &H008080FF&
      Caption         =   "How Many Skirts Do You Need?"
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label lblcostncc 
      BackColor       =   &H008080FF&
      Caption         =   "$38.84 Per Skirt"
      Height          =   255
      Left            =   11160
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblcostflyers 
      BackColor       =   &H008080FF&
      Caption         =   "$44.00 Per Skirt"
      Height          =   255
      Left            =   8160
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblcostflyaway 
      BackColor       =   &H008080FF&
      Caption         =   "$42.50 Per Skirt"
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblcostaline 
      BackColor       =   &H008080FF&
      Caption         =   "$32.99 Per Skirt"
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblncc 
      BackColor       =   &H008080FF&
      Caption         =   "NCC Skirt"
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   13
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblflyers 
      BackColor       =   &H008080FF&
      Caption         =   "Flyers Skirt"
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblflyaway 
      BackColor       =   &H008080FF&
      Caption         =   "Fly Away Skirt"
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblaline 
      BackColor       =   &H008080FF&
      Caption         =   "Aline Skirt"
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblwhichskrit 
      BackColor       =   &H008080FF&
      Caption         =   "Which skirt would you like?"
      BeginProperty Font 
         Name            =   "Dotum"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblskirts 
      BackColor       =   &H008080FF&
      Caption         =   "Now lets look at the skirts!!!!"
      BeginProperty Font 
         Name            =   "NIST Sans Serif"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmSkirts1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Cheerleading (Cheerleading.vbp)
'Form Name : Skirts (Skirts1.frm)
'Author: Heather Johnson
'Date Written: October 28, 2003
'Purpose of Form:'this form will ask you what cheerleading skirt you want
                 'how many skirts you would like to order
                 'your total cost before the discount
                 'your total cost after the discount
Option Explicit
Dim S As Single, Total As Single, HowMany As Integer, J As Single
Private Sub cmdclear_Click()
picSkirttotal.Cls 'clears the skirt total output box
End Sub

Private Sub cmdmainmenu_Click()
frmSkirts1.Hide 'hides the skirts form
frmOrder1.Show 'shows the order form
End Sub

Private Sub cmdquit_Click()
End 'ends the form
End Sub

Private Sub cmdskirttotal_Click()
HowMany = txthowmanyskirts.Text
If optaline = True Then
        picSkirttotal.Print "Your Skirt is the Aline Skirt" 'if you push the optaline button this will print
    ElseIf optflyaway = True Then
        picSkirttotal.Print "Your Skirt is the Fly Away Skirt" 'if you push the optflyaway button this will print
    ElseIf optflyers = True Then
        picSkirttotal.Print "Your Skirt is the Flyers Skirt" 'if you push the optflyers button this is will print
    ElseIf optncc = True Then
        picSkirttotal.Print "Your Skirt is the NCC Skirt" 'if you push the optncc button this will print
End If
If optaline = True Then
    S = 32.99 'if you push the optaline button this is how much each skirt will cost
ElseIf optflyaway = True Then
    S = 42.5  'if you push the optflyaway button this is how much each skirt will cost
ElseIf optflyers = True Then
    S = 44#   'if you push the optflyers button this is how much each skirt will cost
ElseIf optncc = True Then
    S = 38.84 'if you push the optncc button this is how much each skirt will cost
End If
Total = HowMany * S
    picSkirttotal.Print "Your total before the discount for your skirts is "; FormatCurrency(Total, 2)
        'prints out what your total is before you subtract the discount
cmddiscount.Enabled = True 'you can now click on the discount button
End Sub
Private Sub cmddiscount_Click()
Dim Amount(1 To 9) As Integer, Discount(1 To 9) As Single, Found As Boolean, POS As Integer, Savings As Single
picSkirttotal.Print
Found = False
POS = 1
Open PATH & "skirtdiscounts.txt" For Input As #1
For J = 1 To 9
    Input #1, Amount(J), Discount(J)
Next J

Do While Found = False And POS < 9
    POS = POS + 1
    If Total >= Amount(POS) Then
        Found = True
    End If
    
Loop
If Found = True Then
    picSkirttotal.Print "You Save "; FormatPercent(Discount(POS), 0)
Else
    picSkirttotal.Print "Sorry, you don't get a discount"
End If
Close #1
TotalSkirts = Total - (Total * Discount(POS))
picSkirttotal.Print "Your Total with the discount is "; FormatCurrency(TotalSkirts, 2)
End Sub

Private Sub Form_Load()
PATH = "N:\CS130\handin\Johnson_Heather\"
End Sub
