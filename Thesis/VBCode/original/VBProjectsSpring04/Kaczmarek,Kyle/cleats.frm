VERSION 5.00
Begin VB.Form frmcleats 
   BackColor       =   &H00008080&
   Caption         =   "Cleats"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhats 
      BackColor       =   &H0080FF80&
      Caption         =   "Order Hats"
      Height          =   855
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdclearcleats 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear Calculated Price"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdcalccleats 
      BackColor       =   &H0080FF80&
      Caption         =   "Calculate Price of Cleats"
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox txtcleats 
      Height          =   615
      Left            =   3000
      TabIndex        =   13
      Top             =   3960
      Width           =   2415
   End
   Begin VB.PictureBox picresultscleats 
      BackColor       =   &H8000000E&
      Height          =   3375
      Left            =   6000
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   12
      Top             =   3960
      Width           =   4215
   End
   Begin VB.PictureBox picnike 
      Height          =   1575
      Left            =   240
      Picture         =   "cleats.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
   Begin VB.OptionButton optmizuno 
      BackColor       =   &H00008080&
      Caption         =   "Option4"
      Height          =   255
      Left            =   9120
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.OptionButton optreebok 
      BackColor       =   &H00008080&
      Caption         =   "Option3"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.OptionButton optadidas 
      BackColor       =   &H00008080&
      Caption         =   "Option2"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.OptionButton optnike 
      BackColor       =   &H00008080&
      Caption         =   "Option1"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picmizuno 
      Height          =   1575
      Left            =   8040
      Picture         =   "cleats.frx":E80E
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.PictureBox picreebok 
      Height          =   1575
      Left            =   5520
      Picture         =   "cleats.frx":1D01C
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.PictureBox picadidas 
      Height          =   1575
      Left            =   2880
      Picture         =   "cleats.frx":2A0CE
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblmizuno 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "$79.95"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblreebok 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "$75.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6120
      TabIndex        =   20
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lbladidas 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "$94.99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblnike 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "$89.95"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "How Many Pairs of Cleats?"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "Mizuno"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "Reebok"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "Adidas"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "Nike"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "frmcleats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : BaseballUniforms (BaseballUniforms.vbp)
'Form Name : frmcleats (cleats.frm)
'Author: Kyle Kaczmarek
'Date Written: March 15, 2004
'Purpose of Form:'
                 
Private Sub cmdcalccleats_Click()

cleats = txtcleats.Text 'enter how many pairs of cleats you want


If optnike = True Then 'nike cleats are chosen
        cleatsprice = 89.95 'price for nike cleats
    ElseIf optadidas = True Then 'adidas cleats are chosen
        cleatsprice = 94.99 'price for adidas cleats
    ElseIf optreebok = True Then 'reebok cleats are chosen
        cleatsprice = 75.99 'price for reebok cleats
    ElseIf optmizuno = True Then 'mizuno cleats are chosen
        cleatsprice = 79.95 'price for mizuno cleats
End If

cleatstotal = cleats * cleatsprice 'total for cleats
    
picresultscleats.Cls 'clears the picture box
picresultscleats.Print "Pairs of Cleats", "Cost" 'prints out the titles
picresultscleats.Print "***********************", "*****"
picresultscleats.Print Tab(8); cleats, , FormatCurrency(cleatstotal, 2) 'Prints out how many pairs of cleats you ordered and their cost


    
End Sub

Private Sub cmdhats_Click()
frmjerseys.Hide 'closes the jetsey form
frmhats.Show 'shows the hat form
frmorder.Hide 'closes the order form
frmpants.Hide 'closes the pants form
frmcleats.Hide 'closes the cleats form
frmfinal.Hide 'closes the final form
End Sub

