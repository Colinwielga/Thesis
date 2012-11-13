VERSION 5.00
Begin VB.Form Seating 
   BackColor       =   &H80000016&
   Caption         =   "Form9"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   9015
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   5280
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Section 212"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Section 206"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00008000&
      Caption         =   "Section 107"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Section 118"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Finalize Purchase"
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000014&
      Height          =   3495
      Left            =   6600
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "You have 100 seconds to select your seats before your order resets."
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Available Seating"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Seating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St.PaulEvents
'Seating
'David Wutchiett
'February 24, 2010
'User chooses seating section. Prompts credit card input.

Option Explicit

Private Sub Command1_Click()
    Dim cc As String

    Timer1.Enabled = False

    cc = InputBox("Please enter your 10 digit credit card number to complete your purchse")

    Do While Len(cc) <> 10
        cc = InputBox("please enter a 10 digit credit card number")
    Loop

    MsgBox ("Thank you for ordering with St.Paul Events!")

    Form1.Show
    Seating.Hide
End Sub

Private Sub Form_Load()

Timer1.Interval = 1000
Label2.Caption = "20"

End Sub

Private Sub Option1_Click()

If Option1.Value = True Then
    Seatings = "Section 118"
    Option4.Value = False
    Option2.Value = False
    Option3.Value = False
    Command1.Visible = True
    End If
    


End Sub

Private Sub Option2_Click()

If Option2.Value = True Then
    Seatings = "Section 107"
    Option1.Value = False
    Option4.Value = False
    Option3.Value = False
    Command1.Visible = True
    End If
End Sub

Private Sub Option3_Click()

If Option3.Value = True Then
    Seatings = "Section 206"
    Option1.Value = False
    Option2.Value = False
    Option4.Value = False
    Command1.Visible = True
    End If
End Sub

Private Sub Option4_Click()

If Option4.Value = True Then
    Seatings = "Section 212"
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Command1.Visible = True
    End If
    
End Sub

Private Sub Timer1_Timer()
If Label2.Caption = 0 Then
    Timer1.Enabled = False
    MsgBox ("20 secs is up")
    Seating.Hide
    Plans.Show
    
Else
    Label2.Caption = Label2.Caption - 1
End If

End Sub
