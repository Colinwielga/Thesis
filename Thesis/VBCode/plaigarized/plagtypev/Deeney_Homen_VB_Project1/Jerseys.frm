VERSION 5.00
Begin VB.Form Jerseys
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form3"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form3"
   Picture         =   "Jerseys.frx":0000
   ScaleHeight     =   8505
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack
      BackColor       =   &H00E0E0E0&
      Caption         =   "Back <=="
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdNextForm
      BackColor       =   &H00E0E0E0&
      Caption         =   "At this point you have seen some of the top teams and players from around the globe; this leaves only one thing left to do..."
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7200
      Width           =   3375
   End
   Begin VB.CommandButton cmdUnitedStates
      BackColor       =   &H00FF0000&
      Caption         =   "United States"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quit"
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmdPortugal
      BackColor       =   &H0000C000&
      Caption         =   "Portugal"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdItaly
      BackColor       =   &H00FFFF80&
      Caption         =   "Italy"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdBrazil
      BackColor       =   &H0000FFFF&
      Caption         =   "Brazil"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdSpain
      BackColor       =   &H000000C0&
      Caption         =   "Spain"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdFrance
      BackColor       =   &H00800080&
      Caption         =   "France"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton cmdGermany
      BackColor       =   &H000080FF&
      Caption         =   "Germany"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin VB.PictureBox picJersey
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   3360
      ScaleHeight     =   5235
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label lblTeams
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Click on a country's name to see the corresponding team jersey!"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5040
      TabIndex        =   6
      Top             =   6000
      Width           =   1935
   End
End
Attribute VB_Name = "Jerseys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: WorldCup
'Form Name: Jerseys
'Author: Brian Deeney and Nick Homen
'Date written: 2-20-10
'Objective: This is a great form that allows an eye-catching display of the jerseys to be worn by the top World Cup contenders

'Displays Brazil's Jersey
Private Sub cmdBrazil_Click()
picJersey.Cls
picJersey.Picture = LoadPicture(App.Path & "\Brazil_jersey.jpg")
End Sub

'Displays France's Jersey
Private Sub cmdFrance_Click()
picJersey.Cls
picJersey.Cls
picJersey.Picture = LoadPicture(App.Path & "\France_jersey.jpg")
End Sub

'Displays Germany's Jersey
Private Sub cmdGermany_Click()
picJersey.Cls
picJersey.Picture = LoadPicture(App.Path & "\Germany_jersey.jpg")
End Sub

'Allows user to return to statistics form
Private Sub cmdGoBack_Click()
Jerseys.Hide
Stats.Show


End Sub

'Displays Italy's Jersey
Private Sub cmdItaly_Click()
picJersey.Cls
picJersey.Cls
picJersey.Picture = LoadPicture(App.Path & "\Italy_jersey.jpg")
End Sub

'Allows the user to proceed to the World Cup schedule
Private Sub cmdNextForm_Click()
Jerseys.Hide
ComingSoon.Show
End Sub

'Displays Portugal's Jersey
Private Sub cmdPortugal_Click()
picJersey.Cls
picJersey.Picture = LoadPicture(App.Path & "\Portugal_jersey.jpg")
End Sub

'Quit
Private Sub cmdQuit_Click()
End
End Sub
'Displays Spain's Jersey
Private Sub cmdSpain_Click()
picJersey.Cls
picJersey.Cls
picJersey.Picture = LoadPicture(App.Path & "\Spain_jersey.jpg")
End Sub

'Displays United States' Jersey
Private Sub cmdUnitedStates_Click()
picJersey.Cls
picJersey.Picture = LoadPicture(App.Path & "\USA_jersey.jpg")
End Sub

