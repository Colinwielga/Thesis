VERSION 5.00
Begin VB.Form Florence 
   BackColor       =   &H000000FF&
   Caption         =   "Form4"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   15240
   LinkTopic       =   "Form4"
   ScaleHeight     =   10575
   ScaleWidth      =   15240
   Begin VB.PictureBox Picture4 
      Height          =   5655
      Left            =   8280
      Picture         =   "Florence.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   6795
      TabIndex        =   9
      Top             =   5160
      Width           =   6855
   End
   Begin VB.PictureBox Picture3 
      Height          =   3975
      Left            =   10680
      Picture         =   "Florence.frx":14E9A
      ScaleHeight     =   3915
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   4815
      Left            =   1800
      Picture         =   "Florence.frx":191C3
      ScaleHeight     =   4755
      ScaleWidth      =   6555
      TabIndex        =   7
      Top             =   6480
      Width           =   6615
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   0
      Picture         =   "Florence.frx":20937
      ScaleHeight     =   4515
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   2160
      Width           =   3975
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3960
      ScaleHeight     =   2235
      ScaleWidth      =   6675
      TabIndex        =   5
      Top             =   2400
      Width           =   6735
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "What Activities Should I do Then?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtdays 
      Height          =   855
      Left            =   3960
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Pick Another City to Visit Instead"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lbluffizi 
      BackColor       =   &H000000FF&
      Caption         =   "The Uffizi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10680
      TabIndex        =   13
      Top             =   4800
      Width           =   4335
   End
   Begin VB.Label lblmedici 
      BackColor       =   &H000000FF&
      Caption         =   "Medici Chapel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13800
      TabIndex        =   12
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblduomo 
      BackColor       =   &H000000FF&
      Caption         =   "The Duomo and the Bell Tower"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   11
      Top             =   6120
      Width           =   3855
   End
   Begin VB.Label lbldavid 
      BackColor       =   &H000000FF&
      Caption         =   "<-- David at the Accademia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   10
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label lblquestion 
      BackColor       =   &H000000FF&
      Caption         =   "How Many Days Will You Be Staying?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblmain 
      BackColor       =   &H000000FF&
      Caption         =   "What To Do In Florence?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Florence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Where to Travel in Italy
'Form Name: Florence
'Author: Sarah Dayton
'Date Written: March 14, 2009
'This form was made to help the user know what they should do in Florence based on the number of days that they will spend in Florence
Option Explicit
Dim days As Single

Private Sub cmdgo_Click()
days = txtdays
If days <= 1 Then
    picresults.Print "You should go see the Accademia!!"
    picresults.Print "David is HUGE!!!"
    ElseIf days <= 2 Then
        picresults.Print "You should go to: "
        picresults.Print "the Uffizi"
        picresults.Print "and the Accademia!!"
        picresults.Print "It's obviously time to see the art!"
    ElseIf days <= 3 Then
        picresults.Print "You should go to"""
        picresults.Print "the Uffizi,"
        picresults.Print "the Accademia,"
        picresults.Print "and the climb to the top of the Duomo or the Bell Tower!"
        picresults.Print "Get your walking shoes ready!!"
    ElseIf days > 3 Then
        picresults.Print "You should go to"
        picresults.Print "the Uffizi,"
        picresults.Print "the Accademia,"
        picresults.Print "climb both the Bell Tower and the Duomo,"
        picresults.Print "and go to the Medici Chapel!"
        picresults.Print "Here we go!!!"

End If


End Sub

Private Sub cmdgoback_Click()
OpeningPage.Show
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub
