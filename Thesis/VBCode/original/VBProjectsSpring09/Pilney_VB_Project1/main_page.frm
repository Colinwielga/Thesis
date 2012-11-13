VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10035
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7320
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdheavenly 
      Caption         =   "Heavenly, CA/NV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10080
      TabIndex        =   4
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdaspen 
      Caption         =   "Aspen,CO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   3
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdvail 
      Caption         =   "Vail, CO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10080
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdmountsnow 
      Caption         =   "Mount Snow, VT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Select which resort you would like to visit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
   Begin VB.Image Image1 
      Height          =   24585
      Left            =   0
      Picture         =   "main_page.frx":0000
      Top             =   -5760
      Width           =   36870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Ski Trip
'Form Name: main_page
'Author: Sam Pilney
'Written: March 16,2009
'this page allows the user to select which resort they want to visit
'it also brings them to the life ticket form for the corresponding resort
Option Explicit
Dim Continue As Single


Private Sub cmdaspen_Click()
MsgBox "Aspen, a famous Rocky Mountain town in Colorado, is a European-style ski village built in the 1960's and visited by famous celebrities. It is a spectacular world-class ski resort.", , "Aspen and Snowmass Village"
Continue = InputBox("Enter a 1 to continue or a 0 to choose another resort.", "Continue?")
If Continue = 1 Then
    Form1.Hide
    Form2.Show
End If
End Sub

Private Sub cmdheavenly_Click()
MsgBox "Located on Lake Tahoe, Heavenly Ski Resort is the largest ski resort in the US. The resort offers skiing on slopes in both California and Nevada, snowmobiling, sledding, tubing, and a ski school. Heavenly Ski Resort is accessible by a gondola in downtown Lake Tahoe.", , "Heavenly Ski Resort"
Continue = InputBox("Enter a 1 to continue or a 0 to choose another resort.", "Continue?")
If Continue = 1 Then
    Form1.Hide
    Form3.Show
End If
End Sub

Private Sub cmdmountsnow_Click()
MsgBox "Located in Vermont, Mount Snow offers skiing for all abilities. Mount Snow is nestled in the Green Mountains of southern Vermont and for 2 years in a row, was the host of the ESPN 2000 & 2001 Winter X Games.", , "Mount Snow, Vermont"
Continue = InputBox("Enter a 1 to continue or a 0 to choose another resort.", "Continue?")
If Continue = 1 Then
    Form1.Hide
    Form4.Show
End If
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdvail_Click()
MsgBox "A Colorado Rocky Mountain Ski Resort. Vail was born as a European-style ski village in the '60s. This town contributes handsomely to Colorado's colorful reputation. Vail, Colorado, boasts some of the best skiing in the world.", , "Vail: Like Nothing on Earth"
Continue = InputBox("Enter a 1 to continue or a 0 to choose another resort.", "Continue?")
If Continue = 1 Then
    Form1.Hide
    Form5.Show
End If
End Sub
