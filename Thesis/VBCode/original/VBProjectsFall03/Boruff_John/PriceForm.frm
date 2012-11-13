VERSION 5.00
Begin VB.Form PriceForm 
   Caption         =   "Wax Project by John Boruff"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "PriceForm.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReturn1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2415
   End
   Begin VB.PictureBox picMoney 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4200
      ScaleHeight     =   3675
      ScaleWidth      =   7635
      TabIndex        =   3
      Top             =   480
      Width           =   7695
   End
   Begin VB.CommandButton cmdSlow 
      BackColor       =   &H0000FFFF&
      Caption         =   "I don't want to got too fast; or I don't want to spend alot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdMedium 
      BackColor       =   &H0000FFFF&
      Caption         =   "Going Fast can be fun, but it's expensive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdSpeed 
      BackColor       =   &H0000FFFF&
      Caption         =   "Speed at all costs!!!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Priceform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : WaxProject (John Boruff's VB-project.vbp)
'Form Name : PriceForm (PriceForm.frm)
'Author: John Boruff
'purpose of the form:  This form will help the user decide how
                    'much money he'she will want to spen on wax
                    ' by clicking on a price range.

Private Sub cmdMedium_Click() 'informs user about intermediately priced wax
picMoney.Cls  'clears any previous tect in the picture box
picMoney.Print "Mid Range: This wax choice will give better and faster"
picMoney.Print " glide than the “Least Expensive” choice, but at a slightly"
picMoney.Print "higher cost per application. A good choice for everyday skiing,"
picMoney.Print "or race training. "
End Sub

Private Sub cmdReturn1_Click()
    Priceform.Hide  'brings user beck to the MainForm
    MainForm1.Show
End Sub

Private Sub cmdSlow_Click() 'prints text that talks about inexpensive wax
picMoney.Cls  'clears any previous tect in the picture box
picMoney.Print "Least Expensive: This is our recommendation for those"
picMoney.Print "seeking the most economy in their ski wax, and would"
picMoney.Print "great for the recreational Skier and Snowboarder."
End Sub

Private Sub cmdSpeed_Click()  'prints text that tells of expensive wax
picMoney.Cls 'clears any previous tect in the picture box
picMoney.Print "Speed at all costs: This will give the fastest glide possible,"
picMoney.Print " and the best durability. Each layer of wax should be ironed in"
picMoney.Print " and brushed out before the next layer is applied."
picMoney.Print "These waxes are best when speed is crucial, such as race day. "
End Sub

