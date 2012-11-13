VERSION 5.00
Begin VB.Form frmsalsdrink 
   Caption         =   "Drinks"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6960
      TabIndex        =   9
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton cmdleave 
      Caption         =   "Continue on your Tour De St. Joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   7
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdsals 
      Caption         =   "Go Back to Sal's Main "
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   6
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdtequila 
      Caption         =   "Tequila"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   5
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdmixed 
      Caption         =   "Mixed Drinks"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdwine 
      Caption         =   "Wine"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdmartini 
      Caption         =   "Martini"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdshots 
      Caption         =   "Shots"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdbeer 
      Caption         =   "Beer"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "What are you drinking?"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Width           =   5295
   End
   Begin VB.Image Image6 
      Height          =   1800
      Left            =   8280
      Picture         =   "frmsalsdrink.frx":0000
      Top             =   4800
      Width           =   1800
   End
   Begin VB.Image Image5 
      Height          =   1620
      Left            =   8520
      Picture         =   "frmsalsdrink.frx":A902
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Image Image4 
      Height          =   1905
      Left            =   8400
      Picture         =   "frmsalsdrink.frx":10464
      Top             =   720
      Width           =   1305
   End
   Begin VB.Image Image2 
      Height          =   2025
      Left            =   240
      Picture         =   "frmsalsdrink.frx":1879E
      Top             =   4560
      Width           =   1275
   End
   Begin VB.Image Image8 
      Height          =   1680
      Left            =   360
      Picture         =   "frmsalsdrink.frx":20EE0
      Top             =   2760
      Width           =   1110
   End
   Begin VB.Image Image7 
      Height          =   1815
      Left            =   120
      Picture         =   "frmsalsdrink.frx":27122
      Top             =   720
      Width           =   1515
   End
End
Attribute VB_Name = "frmsalsdrink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbeer_Click()

    cmdbeer.Visible = True
    Image7.Visible = True
    cmdshots.Visible = False
    Image8.Visible = False
    cmdmartini.Visible = False
    Image2.Visible = False
    cmdwine.Visible = False
    Image4.Visible = False
    cmdmixed.Visible = False
    Image5.Visible = False
    cmdtequila.Visible = False
    Image6.Visible = False

    Dim number As Single
    number = 0
    
    number = InputBox("How many drinks are you going to have?")
    
    Select Case number
        Case Is <= 3
            MsgBox "Stop at buzzed.  Good for you."
        Case 4 To 6
            MsgBox "You're starting to feel that dancing fever..."
        Case 7 To 11
            MsgBox "With " & number & " beers - it's showing.  You just feel over.  Smooth."
        Case 12 To 14
            MsgBox "You are making a fool out of yourself."
        Case Else
            MsgBox "You should stop drinking.  You dont want to be 'that guy' puking, do you?"
    End Select

End Sub

