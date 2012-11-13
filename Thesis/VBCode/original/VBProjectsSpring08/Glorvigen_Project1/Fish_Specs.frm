VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Fish Journal"
   ClientHeight    =   6465
   ClientLeft      =   4485
   ClientTop       =   3630
   ClientWidth     =   8385
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   8385
   Visible         =   0   'False
   Begin VB.CommandButton cmdj_print 
      Caption         =   "Journal Print"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdleave 
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update Journal"
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
      Left            =   6960
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox picoutput 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   480
      ScaleHeight     =   5355
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   480
      Width           =   6135
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Fisher
'Fish Specs
'Eric Glorvigen
'Date= March 5
Dim ctr As Integer
Dim location(1 To 100) As String
Dim fish(1 To 100) As String
Dim time(1 To 100) As String


Private Sub cmdexit_Click()
    form1.Show
    Form2.Hide
End Sub

Private Sub cmdj_print_Click()
    Dim n As Integer
    
    picoutput.Print "Location:", "Type of Fish:", "Date:"
        picoutput.Print "**************************************************"

    For n = 1 To ctr
                picoutput.Print location(n), fish(n), , time(n)
    Next n
    
End Sub

Private Sub cmdupdate_Click()
    Dim n As Integer
    Dim myfish As String
    Dim spot As String
    Dim mytime As String
       
        Do While spot <> "leave"
            spot = InputBox("Enter Location or type Leave to quit", "Location")
                If spot <> "leave" Then
                    myfish = InputBox("Enter type of fish:", "Fish")
                    mytime = InputBox("Enter date:", "Date")
                    ctr = ctr + 1
                    location(ctr) = spot
                    fish(ctr) = myfish
                    time(ctr) = mytime
                End If
        Loop
End Sub
