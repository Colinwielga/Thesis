VERSION 5.00
Begin VB.Form frmMinnesotaTrees 
   BackColor       =   &H8000000E&
   Caption         =   "Identifying and Organizing Trees of Minnesota"
   ClientHeight    =   6765
   ClientLeft      =   1695
   ClientTop       =   2970
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8610
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "If you already know the kinds of trees you have but want to sort them:                        Click on Tree Below"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   4
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "If you are done with this program then:                 Click on the Tree Below"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Label lblCount 
      BackColor       =   &H8000000E&
      Caption         =   "If already know the kinds of trees you have but want to count them:                            Click on Tree Below"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Label lblGoIdentify 
      BackColor       =   &H8000000E&
      Caption         =   "If you want to try and identify a certain genera of tree of Minnesota:                 Click on Tree Below"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.Image imgSecondTree 
      Height          =   1680
      Left            =   5160
      Picture         =   "frmMinnesotaTrees.frx":0000
      Top             =   1800
      Width           =   2250
   End
   Begin VB.Image imgFirstTree 
      Height          =   1485
      Left            =   1320
      Picture         =   "frmMinnesotaTrees.frx":C602
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Image imgForthtree 
      Height          =   1080
      Left            =   5520
      Picture         =   "frmMinnesotaTrees.frx":160F4
      Top             =   5040
      Width           =   1845
   End
   Begin VB.Image imgThirdTree 
      Height          =   1245
      Left            =   1320
      Picture         =   "frmMinnesotaTrees.frx":1C9D6
      Top             =   5040
      Width           =   1665
   End
   Begin VB.Label lbl 
      BackColor       =   &H8000000E&
      Caption         =   "What kind of trees are in your yard?"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmMinnesotaTrees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Identifying and Organizing sets of Trees from Minnesota
'frmMinnesotaTrees(frmMinnesotaTrees.frm)
'Author: Kelly Fox
'Date Written:3/16/2006
'This purpose of this project is to create an more effiecent method of tree identification, and manipulation of files containing such identifying information
'This is the first form of the project and allows the user to choose what action they wish to take (all other slides have a button to go back to this slide)

Private Sub imgFirstTree_Click()
    frmMinnesotaTrees.Hide
    frmLeaves.Show
End Sub

Private Sub imgForthtree_Click()
    End
End Sub

Private Sub imgSecondTree_Click()
    frmMinnesotaTrees.Visible = False
    frmSort.Visible = True
End Sub

Private Sub imgThirdTree_Click()
    frmMinnesotaTrees.Hide
    frmCount.Show
End Sub
