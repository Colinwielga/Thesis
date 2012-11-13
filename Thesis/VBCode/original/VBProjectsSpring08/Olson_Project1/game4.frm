VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   LinkTopic       =   "Form9"
   Picture         =   "game4.frx":0000
   ScaleHeight     =   9135
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtguess 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   22
      Top             =   2760
      Width           =   2055
   End
   Begin VB.PictureBox picscore 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7680
      ScaleHeight     =   555
      ScaleWidth      =   1755
      TabIndex        =   21
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Next Face!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8040
      Width           =   3375
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   19
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   18
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   16
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   15
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   14
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmd12 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   13
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmd11 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   12
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmd10 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmd18 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   7
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmd17 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   6
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmd16 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmd15 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmd14 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   3
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmd13 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro H"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   2
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdenter 
      Caption         =   "Enter!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdmainmenu 
      Caption         =   "Go to Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Picture         =   "game4.frx":126E5
      TabIndex        =   0
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Ginger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   35
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Gilligan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   34
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Professor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   33
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Mrs Howell"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   32
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Mr Howell"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   31
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Skipper"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   30
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Mary Ann"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   29
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Choices:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   28
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Who is it now??"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   600
      TabIndex        =   27
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Can you recognize a face?"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1200
      TabIndex        =   26
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Type in your guess here:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7440
      TabIndex        =   25
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Score so far:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7680
      TabIndex        =   24
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label lblanswer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   23
      Top             =   5640
      Width           =   1935
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim score As Single


Private Sub cmd1_Click()
    cmd1.Visible = False
    If cmd1.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd10_Click()
    cmd10.Visible = False
    If cmd10.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd11_Click()
    cmd11.Visible = False
    If cmd11.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd12_Click()
    cmd12.Visible = False
    If cmd12.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd13_Click()
    cmd13.Visible = False
    If cmd13.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd14_Click()
    cmd14.Visible = False
    If cmd14.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd15_Click()
    cmd15.Visible = False
    If cmd15.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd16_Click()
    cmd16.Visible = False
    If cmd16.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd17_Click()
    cmd17.Visible = False
    If cmd17.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd18_Click()
    cmd18.Visible = False
    If cmd18.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd2_Click()
    cmd2.Visible = False
    If cmd2.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd3_Click()
    cmd3.Visible = False
    If cmd3.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd4_Click()
    cmd4.Visible = False
    If cmd4.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd5_Click()
    cmd5.Visible = False
    If cmd5.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd6_Click()
    cmd6.Visible = False
    If cmd6.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd7_Click()
    cmd7.Visible = False
    If cmd7.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd8_Click()
    cmd8.Visible = False
    If cmd8.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmd9_Click()
    cmd9.Visible = False
    If cmd9.Visible = False Then
        score = score + 10
    Else: score = score
    End If
End Sub

Private Sub cmdenter_Click()
    If LCase(Trim(txtguess.Text)) = LCase("Mr Howell") Then
        lblanswer = "Correct!!"
        score = 180 - score
        totalscore = totalscore + score
        picscore.Print totalscore
        picscore.Print "out of a possible 720."
    Else: lblanswer = "Try again!"
        score = score + 10
    End If
    

End Sub

Private Sub cmdmainmenu_Click()
    Form9.Hide
    Form1.Show
End Sub


Private Sub cmdnext_Click()
    Form9.Hide
    Form10.Show
End Sub
