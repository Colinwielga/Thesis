VERSION 5.00
Begin VB.Form frmwhichfact 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "What Would You Like To Know About The Olympic 100 Meter Dash?"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdyourcountry 
      BackColor       =   &H00008000&
      Caption         =   "How Has Your Country Fared?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmwhichfact.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdtrend 
      BackColor       =   &H00008000&
      Caption         =   "What Time Trends Show Through History?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   240
      Picture         =   "frmwhichfact.frx":1E637
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdcountrywins 
      BackColor       =   &H00008000&
      Caption         =   "Which Countries Are Historically The Best?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   8880
      Picture         =   "frmwhichfact.frx":2D46E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdaverage 
      BackColor       =   &H00008000&
      Caption         =   "What Is The Overall Average Times Of All Olympic 100 Meter Runners?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   6720
      Picture         =   "frmwhichfact.frx":2F86B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdfastest 
      BackColor       =   &H00008000&
      Caption         =   "Who's The Fastest?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   2400
      Picture         =   "frmwhichfact.frx":31D5D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lbldirection 
      BackColor       =   &H0000FFFF&
      Caption         =   "                       Click on Any Photo Above To Begin"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   7680
      Width           =   10815
   End
End
Attribute VB_Name = "frmwhichfact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaverage_Click()
    frmwhichfact.Hide
    frmaverage.Show
End Sub

Private Sub cmdcountrywins_Click()
    frmwhichfact.Hide
    frmbestcountries.Show
End Sub

Private Sub cmdfastest_Click()
     frmwhichfact.Hide
     frmcompare.Show
End Sub

Private Sub cmdtrend_Click()
    frmwhichfact.Hide
    frmtimetrends.Show
End Sub

Private Sub cmdyourcountry_Click()
    frmwhichfact.Hide
    frmyourcountry.Show
End Sub
