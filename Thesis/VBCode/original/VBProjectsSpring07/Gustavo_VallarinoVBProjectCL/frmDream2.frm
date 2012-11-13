VERSION 5.00
Begin VB.Form frmDream 
   BackColor       =   &H8000000D&
   Caption         =   "Dream Team"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfinaldream 
      Caption         =   "See your Dream Team"
      Height          =   735
      Left            =   3720
      TabIndex        =   26
      Top             =   7440
      Width           =   2895
   End
   Begin VB.ComboBox Comborightfwd 
      Height          =   315
      Left            =   6240
      TabIndex        =   25
      Text            =   "Right Forward"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.ComboBox comboLeftfwd 
      Height          =   315
      Left            =   3120
      TabIndex        =   24
      Text            =   "Left Forward"
      Top             =   6480
      Width           =   2055
   End
   Begin VB.ComboBox combooffensive 
      Height          =   315
      Left            =   8520
      TabIndex        =   23
      Text            =   "Offensive Mid"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ComboBox comboleftmid 
      Height          =   315
      Left            =   5400
      TabIndex        =   22
      Text            =   "Left Mid"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ComboBox comborightmid 
      Height          =   315
      Left            =   2880
      TabIndex        =   21
      Text            =   "Right Mid"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.ComboBox combobackmid 
      Height          =   315
      Left            =   360
      TabIndex        =   20
      Text            =   "Back Mid"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox ComboCenterD2 
      Height          =   315
      Left            =   9240
      TabIndex        =   19
      Text            =   "Central Defender"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox ComboCenterD1 
      Height          =   315
      Left            =   6840
      TabIndex        =   18
      Text            =   "Central Defender"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ComboBox ComboLd 
      Height          =   315
      Left            =   4560
      TabIndex        =   17
      Text            =   "Left Defender"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox comboRightD 
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      Text            =   "Right Defender"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox combogoalie 
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Text            =   "Goalie"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select "
      Height          =   1455
      Left            =   240
      TabIndex        =   12
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Back to Main Menu"
      Height          =   735
      Left            =   8520
      TabIndex        =   11
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label lbldfCnt2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Central Defender 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   9240
      TabIndex        =   15
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblselect 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Press To Select your Dream Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   975
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblfwd2 
      BackColor       =   &H8000000D&
      Caption         =   "Right Forward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   6000
      TabIndex        =   10
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label lblfwd1 
      BackColor       =   &H8000000D&
      Caption         =   "Left Forward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lblmidfw 
      BackColor       =   &H8000000D&
      Caption         =   "Offensive Midfield"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   8400
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblmidlft 
      BackColor       =   &H8000000D&
      Caption         =   "Left Midfield"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblmidrgt 
      BackColor       =   &H8000000D&
      Caption         =   "Right Midfield"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   2760
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblmidbk 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Back Mid "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblCentral 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Central Defender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblleft 
      BackColor       =   &H8000000D&
      Caption         =   "Left Defender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Right Defender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblGoalie 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Goalie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblChose 
      BackColor       =   &H8000000D&
      Caption         =   "Choose Your Dream Team"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmDream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this page was to allow me to choose from a set of players that would make up
'a dream team, and later they would be printed in another page
'I used combo boxes so the first thing i needed to do was to add the items to the combobox and name them appropiatly
'then i used the IF/The/Elseif conditions to select the one that would be printed and what would they equal


Private Sub cmdfinaldream_Click()
frmDream.Hide
frmDreamShow.Show

End Sub

Private Sub cmdMain_Click()
frmChampions.Show
frmDream.Hide
End Sub

Private Sub cmdSelect_Click()
'Used the list index according to the position in the combobox to assign them to a variable
'So if they according to what variable the user would choose it would save and assign variable

If combogoalie.ListIndex = 0 Then
Goalie = "Oliver Kahn"
ElseIf combogoalie.ListIndex = 1 Then
Goalie = "Iker Casillas"
End If

If comboRightD.ListIndex = 0 Then
RightD = "Eric Abidal"
ElseIf combogoalie.ListIndex = 1 Then
RightD = "Paolo Maldini"
End If

If ComboLd.ListIndex = 0 Then
LeftD = "Roberto Carlos"
ElseIf ComboLd.ListIndex = 1 Then
LeftD = "Lucio"
End If

If ComboCenterD1.ListIndex = 0 Then
Centerd1 = "Fabio Cannavaro"
ElseIf ComboCenterD1.ListIndex = 1 Then
Centerd1 = "Rio Ferdinand"
End If

If ComboCenterD2.ListIndex = 0 Then
centerD2 = "John Terry"
ElseIf ComboCenterD1.ListIndex = 1 Then
centerD2 = "Sergio Ramos"
End If

If combobackmid.ListIndex = 0 Then
backmid = "Frank Lampard"
ElseIf combobackmid.ListIndex = 1 Then
backmid = "Genaro Gattuso"
End If

If comborightmid.ListIndex = 0 Then
Rightmid = "Kaka"
ElseIf comborightmid.ListIndex = 1 Then
Rightmid = "Malouda"
End If

If comboleftmid.ListIndex = 0 Then
leftmid = "Cristiano Ronaldo"
ElseIf comboleftmid.ListIndex = 1 Then
leftmid = "Robinho"
End If

If combooffensive.ListIndex = 0 Then
offensive = "Ronaldinho"
ElseIf ComboCenterD1.ListIndex = 1 Then
offensive = "Zinedine Zidane"
End If

If comboLeftfwd.ListIndex = 0 Then
leftfwd = "Ronaldo"
ElseIf comboLeftfwd.ListIndex = 1 Then
leftfwd = "Andriy Shevchenko"
End If

If Comborightfwd.ListIndex = 0 Then
rightfwd = "Raul"
ElseIf Comborightfwd.ListIndex = 1 Then
rightfwd = "Henry"
End If


End Sub

'With this funcion AddItem i added items to my combo boxes
'In this part i had to be very careful namimg each combo box

Private Sub Form_Load()
combogoalie.AddItem ("Oliver Kahn")
combogoalie.AddItem ("Iker Casillas")

comboRightD.AddItem ("Eric Abidal")
comboRightD.AddItem ("Paolo Maldini")

ComboLd.AddItem ("Roberto Carlos")
ComboLd.AddItem ("Lucio")

ComboCenterD1.AddItem ("Fabio Canavaro")
ComboCenterD1.AddItem ("Rio Ferdinand")

ComboCenterD2.AddItem ("John Terry")
ComboCenterD2.AddItem ("Sergio Ramos")

combobackmid.AddItem ("Frank Lampard")
combobackmid.AddItem ("Genaro Gattuso")

comborightmid.AddItem ("Kaka")
comborightmid.AddItem ("Malouda")

comboleftmid.AddItem ("Cristiano Ronaldo")
comboleftmid.AddItem ("Robinho")

combooffensive.AddItem ("Ronaldinho")
combooffensive.AddItem ("Zinedine Zidane")

comboLeftfwd.AddItem ("Ronaldo")
comboLeftfwd.AddItem ("Andriy Shevchenko")

Comborightfwd.AddItem ("Raul")
Comborightfwd.AddItem ("Henry")


End Sub
