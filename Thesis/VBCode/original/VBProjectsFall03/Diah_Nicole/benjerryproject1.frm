VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   600
   ClientTop       =   990
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11415
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF00FF&
      Height          =   1095
      Left            =   1200
      Picture         =   "benjerryproject1.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF00FF&
      Height          =   1215
      Left            =   9000
      Picture         =   "benjerryproject1.frx":1137
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF00FF&
      Height          =   975
      Left            =   5160
      Picture         =   "benjerryproject1.frx":1B74
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   240
      Picture         =   "benjerryproject1.frx":28EC
      ScaleHeight     =   1155
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdtourform 
      Caption         =   "Tour Info"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.PictureBox picdesign2 
      BackColor       =   &H00FF00FF&
      Height          =   975
      Left            =   240
      Picture         =   "benjerryproject1.frx":3D56
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox picdesign1 
      BackColor       =   &H00FF00FF&
      Height          =   975
      Left            =   10080
      Picture         =   "benjerryproject1.frx":47B1
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.PictureBox picturetitle 
      BackColor       =   &H00FF00FF&
      Height          =   1215
      Left            =   2880
      Picture         =   "benjerryproject1.frx":505E
      ScaleHeight     =   1155
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   840
      Width           =   5655
   End
   Begin VB.CommandButton cmdnutritionform 
      Caption         =   "Nutritional value"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Exit"
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
      Left            =   9240
      TabIndex        =   1
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdflavorsform 
      Caption         =   "Rate your Favorite Flavors"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF00FF&
      Caption         =   "Nicole Diah  CS130"
      Height          =   495
      Left            =   9120
      TabIndex        =   17
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF00FF&
      Caption         =   "Ben and Jerry's is proud to be a part of One Sweet Whirled in one sweet campaign to fight global warming."
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   7320
      Width           =   4455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Ice Cream Factory"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   $"benjerryproject1.frx":6BF2
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8640
      TabIndex        =   10
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   $"benjerryproject1.frx":6C8C
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4800
      TabIndex        =   9
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "Do you know what's in your ice cream? Find out here."
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4260
      TabIndex        =   6
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name : benjerryproject.vbp (Nicole Diah's VB Project.vbp)
'Form Name : form 1 (benjerryproject1.frm)
'Author: Nicole Diah
'Date Written: Oct. 27, 2003
'Purpose of Form: Used as a main page to connect all of the
                ' other forms
                ' This program is designed to give users more information
                ' about various features pertaining to the Ben & Jerry's company.

Private Sub cmdflavorsform_Click()
Form1.Hide
flavorsform.Show
nutritionform.Hide
toursform.Hide
End Sub

Private Sub cmdnutritionform_Click()
Form1.Hide
flavorsform.Hide
nutritionform.Show
toursform.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdtourform_Click()
Form1.Hide
flavorsform.Hide
nutritionform.Hide
toursform.Show
End Sub
