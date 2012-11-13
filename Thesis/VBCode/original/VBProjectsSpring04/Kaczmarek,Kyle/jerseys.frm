VERSION 5.00
Begin VB.Form frmjerseys 
   BackColor       =   &H0080FFFF&
   Caption         =   "Jerseys"
   ClientHeight    =   9090
   ClientLeft      =   480
   ClientTop       =   1305
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   14235
   Begin VB.CommandButton cmdpants 
      BackColor       =   &H0080C0FF&
      Caption         =   "Order Pants"
      Height          =   735
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8040
      Width           =   2775
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H0080C0FF&
      Caption         =   "Clear Calculated Price of Jerseys"
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8040
      Width           =   3015
   End
   Begin VB.PictureBox picjerseycost 
      BackColor       =   &H8000000E&
      Height          =   2775
      Left            =   8640
      ScaleHeight     =   2715
      ScaleWidth      =   4755
      TabIndex        =   19
      Top             =   5520
      Width           =   4815
   End
   Begin VB.CommandButton cmdjerseycost 
      BackColor       =   &H0080C0FF&
      Caption         =   "Calculate Cost of Jerseys"
      Height          =   855
      Left            =   3000
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6960
      Width           =   2655
   End
   Begin VB.TextBox txtjerseys 
      Height          =   615
      Left            =   4080
      TabIndex        =   16
      Top             =   6120
      Width           =   2655
   End
   Begin VB.OptionButton optreds 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   12720
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picreds 
      Height          =   2535
      Left            =   11640
      Picture         =   "jerseys.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.OptionButton optyankees 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   9960
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton opttwins 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optphillies 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optsox 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picyankees 
      Height          =   2535
      Left            =   8760
      Picture         =   "jerseys.frx":16B42
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.PictureBox pictwins 
      Height          =   2535
      Left            =   5880
      Picture         =   "jerseys.frx":2D684
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.PictureBox picphillies 
      Height          =   2535
      Left            =   3000
      Picture         =   "jerseys.frx":441C6
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.PictureBox picsox 
      Height          =   2535
      Left            =   120
      Picture         =   "jerseys.frx":5AD08
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "$45 per Jersey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11760
      TabIndex        =   26
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "$40 per Jersey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   9120
      TabIndex        =   25
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "$85 per Jersey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   6240
      TabIndex        =   24
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "$50 per Jersey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "$75 per Jersey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "How Many Jerseys?"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   17
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Choose A Jersey!!"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   3360
      TabIndex        =   15
      Top             =   360
      Width           =   8535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Sleeveless"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11640
      TabIndex        =   14
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Team Logo"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Retro"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Pinstripe"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Team Nickname With Trim"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   2655
   End
End
Attribute VB_Name = "frmjerseys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : BaseballUniforms (BaseballUniforms.vbp)
'Form Name : frmjerseys (jerseys.frm)
'Author: Kyle Kaczmarek
'Date Written: March 15, 2004
'Purpose of Form:'

Private Sub cmdclear_Click()

picjerseycost.Cls 'clears the picture box
txtjerseys = "" 'puts nothing in the input box


End Sub

Private Sub cmdjerseycost_Click()


jerseys = txtjerseys.Text 'enter how many jerseys you want


If optsox = True Then 'chosen sox option
        jerseysprice = 75 'sox price
    ElseIf optphillies = True Then 'chosen phillies option
        jerseysprice = 50 'phillies price
    ElseIf opttwins = True Then 'chosen twins price
        jerseysprice = 85 'twins price
    ElseIf optyankees = True Then 'chosen yankees option
        jerseysprice = 40 'yankees price
    ElseIf optreds = True Then 'chosen reds option
        jerseysprice = 45 'reds price
End If

jerseytotal = jerseys * jerseysprice 'multiplies the total number of jerseys needed by the price
    
picjerseycost.Cls 'clears the picture box
picjerseycost.Print "Number of Jerseys", "Cost" 'prints out the titles
picjerseycost.Print "***********************", "*****" 'prints out the stars
picjerseycost.Print Tab(8); jerseys, , FormatCurrency(jerseytotal, 2) 'prints out the totals
    



End Sub

Private Sub cmdpants_Click()
frmjerseys.Hide 'closes the jersey form
frmhats.Hide 'closes the hat form
frmorder.Hide 'closes the order form
frmpants.Show 'shows the pants form
frmcleats.Hide 'closes the cleats form
frmfinal.Hide 'closes the final form
End Sub

