VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Weight Calculator"
   ClientHeight    =   6090
   ClientLeft      =   3780
   ClientTop       =   2700
   ClientWidth     =   7785
   LinkTopic       =   "Form8"
   ScaleHeight     =   6090
   ScaleWidth      =   7785
   Visible         =   0   'False
   Begin VB.CommandButton cmdmove 
      BackColor       =   &H0080C0FF&
      Caption         =   "See if you have a record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.CheckBox Checksteelhead 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Steelhead"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H008080FF&
      Caption         =   "Leave Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   975
   End
   Begin VB.CheckBox Checksturgeon 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sturgeon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox Checkperch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Perch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CheckBox Checknorthern 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nothern Pike"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox Checkbluegill 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bluegill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CheckBox Checksmall 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Smallmouth Bass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CheckBox Checkwalleye 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Walleye"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CheckBox Checkmusky 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Musky"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find Weight"
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
      Left            =   2880
      Picture         =   "Weight_Calculator.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CheckBox checkcrappie 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crappie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0080FF80&
      Caption         =   "Back To Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Please Check Box To Find Weight of Fish (only check one box at a time)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Minnesota Fisher
'Weight Calculator
'Eric Glorvigen
'Date= March 5
' this page uses the option buttons and asks the user for input through input boxes
' and gives the user feedback through a msgbox, there is also a little joke inside


Private Sub cmdexit_Click()
    'returns to main pare
        form1.Show
        Form8.Hide
End Sub

Private Sub cmdfind_Click()
'finds the weight of the fish which is checked, there is an if statement for
'each check box, and if it is 1 then there are certain input boxs to ask the user
'for two variables then it calculates them to find the weight of the fish


    Dim inch As Single, girth As Single
    
    If checkcrappie.Value = 1 Then
       inch = InputBox("Please Enter Length in Inches:", "Crappie")
        girth = InputBox("Please Enter the Girth in Inches:", "Crappie")
        weightcrappie = girth * girth * (inch / 800)
        MsgBox "The weight of your Crappie roughly weighs " & FormatNumber(weightcrappie, 2) & " pounds", , "Weight"
    End If
    
    If Checkmusky.Value = 1 Then
        inch = InputBox("Please Enter Length in Inches:", "Musky")
        girth = InputBox("Please Enter the Girth in Inches:", "Musky")
        weightmusky = girth * girth * (inch / 800)
        MsgBox "The weight of your musky roughly weighs " & FormatNumber(weightmusky, 2) & " pounds", , "Weight"
    End If
    
    
    If Checkwalleye.Value = 1 Then
       inch = InputBox("Please Enter Length in Inches:", "Walleye")
        girth = InputBox("Please Enter the Girth in Inches:", "Walleye")
        weightwalleye = girth * girth * (inch / 800)
        MsgBox "The weight of your Walleye roughly weighs " & FormatNumber(weightwalleye, 2) & " pounds", , "Weight"
    End If
    
    
    If Checksmall.Value = 1 Then
       inch = InputBox("Please Enter Length in Inches:", "Smallmouth")
        girth = InputBox("Please Enter the Girth in Inches:", "Smallmouth")
        weightsmall = girth * girth * (inch / 800)
        MsgBox "The weight of your Smallmouth roughly weighs " & FormatNumber(weightsmall, 2) & " pounds", , "Weight"
    End If
    
    
    If Checkbluegill.Value = 1 Then
       inch = InputBox("Please Enter Length in Inches:", "Bluegill")
        girth = InputBox("Please Enter the Girth in Inches:", "Bluegill")
        weightbluegill = girth * girth * (inch / 800)
        MsgBox "The weight of your Bluegill roughly weighs " & FormatNumber(weightbluegill, 2) & " pounds", , "Weight"
    End If
    
    
    If Checknorthern.Value = 1 Then
       inch = InputBox("Please Enter Length in Inches:", "Northern")
        girth = InputBox("Please Enter the Girth in Inches:", "Nothern")
        weightnorthern = girth * girth * (inch / 800)
        MsgBox "The weight of your Northern roughly weighs " & FormatNumber(weightnorthern, 2) & " pounds", , "Weight"
    End If
    
    
    If Checkperch.Value = 1 Then
        MsgBox "Are you Really trying to find the weight of a perch?", , "Come On " & inputname & "!"
    End If
    
    
    If Checksturgeon.Value = 1 Then
       inch = InputBox("Please Enter Length in Inches:", "Sturgeon")
        girth = InputBox("Please Enter the Girth in Inches:", "Sturgeon")
        weightsturgeon = girth * girth * (inch / 800)
        MsgBox "The weight of your Sturgeon roughly weighs " & FormatNumber(weightsturgeon, 2) & " pounds", , "Weight"
    End If
    
    
    If Checksteelhead.Value = 1 Then
       inch = InputBox("Please Enter Length in Inches:", "Steelhead")
        girth = InputBox("Please Enter the Girth in Inches:", "Steelhead")
        weightsteelhead = girth * girth * (inch / 800)
        MsgBox "The weight of your Steelhead roughly weighs " & FormatNumber(weightsteelhead, 2) & " pounds", , "Weight"
    End If
    
    
    
        
    
   
End Sub

Private Sub cmdmove_Click()
    'this button shows the state record page
        Form8.Hide
        Form5.Show
End Sub

Private Sub cmdquit_Click()
    'exits the program
        End
End Sub

