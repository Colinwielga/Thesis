VERSION 5.00
Begin VB.Form Trout 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton GotoStreams 
      BackColor       =   &H000080FF&
      Caption         =   "Let's Find Some Trout Streams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7680
      Width           =   2295
   End
   Begin VB.CommandButton Quit3 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox TextBox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   6
      Top             =   7680
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0FF&
      Height          =   2175
      Left            =   3120
      Picture         =   "Trout.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   5160
      Width           =   5655
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   3120
      Picture         =   "Trout.frx":4530
      ScaleHeight     =   2115
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   2760
      Width           =   5655
   End
   Begin VB.PictureBox picBrown 
      Height          =   2175
      Left            =   3120
      Picture         =   "Trout.frx":8A40
      ScaleHeight     =   2115
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   $"Trout.frx":CE53
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Rainbow Trout "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Brook Trout "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Brown Trout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Trout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is designed to give the user a visual reference about what each species of Trout in
'Minnesota looks like and also to give them a guide about how big they can get.
'Instead of using command buttons, this form uses clickable picture boxes to execute the action.
'It makes the button both functional and interesting to look at.
'
'Kevin Lundeen
'March 23rd
'

Dim Length As Single

Private Sub GotoStreams_Click()
    Streams.Show
End Sub

'This subroutine takes a number from a text box and compares it to the State record Brown Trout in Minnesota
'It uses the select case function to accomplish this

Private Sub picBrown_Click()

    Length = TextBox
    
    Select Case Length
        Case Is > 31.4
            MsgBox ("Too big.")
        Case Is > 31.4
            MsgBox ("Just a Little smaller now.")
        Case Is = 31
            MsgBox ("You got it.  31.4 Inches to be specific.")
        Case Is > 30
            MsgBox ("Pretty close. But still too small.")
        Case Is > 25
            MsgBox ("Close but still too small.")
        Case Is > 20
            MsgBox ("Not nearly big enough.")
        Case Is > 15
            MsgBox ("Keep on guessing.")
        Case Is <= 15
            MsgBox ("It's a lot bigger.")
        Case Else
            MsgBox ("Do you know anything about fishing?")
    End Select
End Sub


'This subroutine uses data from an input box that it compares with the state record brook trout
'It uses the select case function to accomplish this

Private Sub Picture2_Click()

    Length = TextBox
        
    Select Case Length
        Case Is > 30
            MsgBox ("Way too big.")
        Case Is > 24
            MsgBox ("A little bit smaller now.")
        Case Is = 24
            MsgBox ("You got it.  24 Inches exactly.")
        Case Is > 20
            MsgBox ("Pretty close. But still too small.")
        Case Is > 15
            MsgBox ("That's not even a trophy fish, keep guessing.")
        Case Is > 10
            MsgBox ("Maybe you should try doubling that length.")
        Case Else
            MsgBox ("Do you know anything about fishing?")
    End Select

End Sub


'This subroutine uses data from an input that it compares with the state record rainbow trout
'It uses the select case function to accomplish this

Private Sub Picture3_Click()
    
    Length = TextBox
    
    Select Case Length
        Case Is > 40
            MsgBox ("You're getting a little too big. I'd try around 33.")
        Case Is > 33
            MsgBox ("Too big.")
        Case Is = 33
            MsgBox ("You got it.  33 Inches exactly.")
        Case Is > 30
            MsgBox ("Pretty close. But still too small.")
        Case Is > 25
            MsgBox ("Getting Closer. But go bigger.")
        Case Is > 20
            MsgBox ("Not nearly big enough.")
        Case Is > 15
            MsgBox ("Keep on guessing.")
        Case Is <= 15
            MsgBox ("It's a lot bigger.")
        Case Else
            MsgBox ("Do you know anything about fishing?")
    End Select
    
End Sub

Private Sub Quit3_Click()
    End
End Sub
