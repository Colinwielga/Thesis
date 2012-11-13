VERSION 5.00
Begin VB.Form FrmEverythingElseInBetweenMain 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12885
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHappy 
      BackColor       =   &H00FF00FF&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox BubbleBath 
      BackColor       =   &H0000FFFF&
      Height          =   2175
      Left            =   10320
      Picture         =   "EverythingElseInBetweenMain.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   10
      Top             =   6240
      Width           =   2175
   End
   Begin VB.PictureBox SleepingCat 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   240
      Picture         =   "EverythingElseInBetweenMain.frx":0C0D
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdHygiene 
      BackColor       =   &H00FF00FF&
      Caption         =   "Why is good Hygiene necessary?"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   2775
   End
   Begin VB.CommandButton cmdSleep 
      BackColor       =   &H00FF00FF&
      Caption         =   "Do you get enough Sleep?"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox txtHappy 
      Height          =   975
      Left            =   8760
      TabIndex        =   6
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturntoMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00FF0000&
      Caption         =   "Click below to find out if you 've been getting enough                  ZZZZ's and are squeaky clean!"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   5400
      Width           =   7335
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00FF0000&
      Caption         =   "Two other very important components of health are: SLEEP and HYGIENE"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   5040
      Width           =   10335
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Here :"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00FF0000&
      Caption         =   "On a scale from 1 to 5, How happy are you?"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1095
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   5895
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FF0000&
      Caption         =   " ...because there is more to health than just food and exercise!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   10575
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FF0000&
      Caption         =   "Everything Else    In Between..."
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   2655
      Left            =   3480
      TabIndex        =   2
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "FrmEverythingElseInBetweenMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Bennie Health Project
    'FrmEverythingElseInBetween
    'Heidi Donnelly
    'Written: 10/5
    'The purpose of this form is to provide the user with other everyday areas that are necessary to look into in order to be healthy such as Sleep and Hygiene.
    
Private Sub cmdHappy_Click()
'this button displays a response to the number from 1 to 5 that the user inputed regarding their happiness level.
Dim Happiness As Integer
Dim Results As String
    
'intialize variables
Happiness = txtHappy.Text

Select Case Happiness
        Case Is >= 5
            MsgBox ("I am glad to hear that you are very happy, ") & UserName & (". Happiness brings health!")
        Case 4
            MsgBox ("That is great! Sometimes life isn't always perfect but you can still be happy!")
        Case 3, 2
            MsgBox ("Sometimes you need to take some time out for yourself and really do something that you love. Have some ") & UserName & (" time. This will help lift your mood!")
        Case 1
            MsgBox ("Everyone has bad days. If this is a continually thing, it may help to talk to some one about how you're feeling!")
        Case 0
            MsgBox ("Depression is a very common thing among women. I recommend that you speak to your doctor soon about your unhappiness. It helps to talk to someone about how you're feeling. You cannot be healthy if you are unhappy!")
        Case Else
            MsgBox ("Oops, invalid number! Try again!")
End Select

End Sub

'this button will simply display a message regarding hygiene
Private Sub cmdHygiene_Click()
MsgBox ("No matter how attractive your clothing may be, you can not be well-dressed unless your body is also well cared for. No matter how healthy you are inside, you won't look healthy unless your outside is healthy too. This means that the hair, skin, teeth, and hands all contribute to a well groomed, healthy appearance.")
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturntoMain_Click()
    FrmEverythingElseInBetweenMain.Hide
    FrmMain.Show
End Sub

Private Sub cmdSleep_Click()
    FrmSleep.Show
    FrmEverythingElseInBetweenMain.Hide
End Sub
