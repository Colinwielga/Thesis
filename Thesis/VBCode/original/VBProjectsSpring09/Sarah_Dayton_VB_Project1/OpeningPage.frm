VERSION 5.00
Begin VB.Form OpeningPage 
   Caption         =   "Form1"
   ClientHeight    =   10575
   ClientLeft      =   1950
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   Picture         =   "OpeningPage.frx":0000
   ScaleHeight     =   10575
   ScaleWidth      =   10605
   Begin VB.CommandButton cmdslideshow 
      Caption         =   "What Does Italy Look Like?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdnaples 
      Caption         =   "Naples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdrome 
      Caption         =   "Rome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdflorence 
      Caption         =   "Florence"
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
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdvenice 
      Caption         =   "Venice"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdmilan 
      BackColor       =   &H00C00000&
      Caption         =   "Milan"
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
      Left            =   1800
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Click On The City That You Want To Visit!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   9600
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "Where Do You Want to Go In Italy?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   7440
      Width           =   4335
   End
End
Attribute VB_Name = "OpeningPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Where to Travel in Italy
'Form Name: Opening Page
'Author: Sarah Dayton
'This project is to help the user decide what they would like to do in Italy, where they should travel, and what they can see
Option Explicit

Private Sub cmdflorence_Click()
OpeningPage.Hide
Milan.Hide
Venice.Hide
Florence.Show
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub

Private Sub cmdmilan_Click()
OpeningPage.Hide
Milan.Show
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub

Private Sub cmdnaples_Click()
OpeningPage.Hide
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Show
SlideShowItaly.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdrome_Click()
OpeningPage.Hide
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Show
Naples.Hide
SlideShowItaly.Hide
End Sub

Private Sub cmdslideshow_Click()
OpeningPage.Hide
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Show
End Sub

Private Sub cmdvenice_Click()
OpeningPage.Hide
Milan.Hide
Venice.Show
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub
