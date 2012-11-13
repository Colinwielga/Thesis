VERSION 5.00
Begin VB.Form frmchat 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "MiM:"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6615
   Begin VB.TextBox txttyped 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   6630
   End
   Begin VB.PictureBox picconvo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   6645
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   6640
      Begin VB.Timer Time2 
         Interval        =   1
         Left            =   4440
         Top             =   2040
      End
      Begin VB.Label lblheader 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   6645
      End
   End
   Begin VB.Label lblabouttext 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4810
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label lblopen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "o:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   4900
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblwink 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(;"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   4900
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblsad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "):"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   4900
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblhappy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4900
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblabout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "about: MiM"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblinsert 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "insert (:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblsend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "send"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblsep 
      BackColor       =   &H00DA6903&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   6640
   End
End
Attribute VB_Name = "frmchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim click As Integer
    Dim click2 As Integer
Private Sub cmdsend_Click()
    Dim typed As String
    typed = "Guest: " & txttyped.Text
    picconvo.ForeColor = &H800000
    c = c
    If c > 12 Then
        picconvo.top = picconvo.top - 390
        picconvo.top = picconvo.top
        picconvo.Height = picconvo.Height + 390
        picconvo.Height = picconvo.Height
        lblheader.top = lblheader.top + 390
        lblheader.top = lblheader.top
    End If
    picconvo.Print typed
    picconvo.Print
    c = c + 2
    Open App.Path & "\guestsaid.txt" For Output As #1
        Print #1, typed
        Close #1
    txttyped.Text = ""
    txttyped.SetFocus
    picconvo.ForeColor = &H8080&
End Sub

Private Sub Form_Load()
    picconvo.Cls
    picconvo.Height = 3135
    picconvo.top = 0
    lblheader.top = 0
    c = 0
    lblheader.Caption = "MiM: chat with " & buddy
    picconvo.Print
    picconvo.Print
    txttyped.Enabled = True
End Sub

Private Sub lblabout_Click()
    click = click
        If click = 0 Then
            frmchat.Height = 5775
            lblabouttext.Visible = True
            lblabouttext.Caption = "MiM: was created by Richard Motzko, 2005"
            click = 1
        Else
            lblabouttext.Visible = False
            frmchat.Height = 5310
            click = 0
        End If
End Sub

Private Sub lblhappy_Click()
    txttyped.Text = txttyped.Text & " ( :"
    lblhappy.Visible = False
    lblopen.Visible = False
    lblsad.Visible = False
    lblwink.Visible = False
    frmchat.Height = 5310

End Sub

Private Sub lblinsert_Click()
click2 = click2
    If click2 = 0 Then
        frmchat.Height = 5775
        lblhappy.Visible = True
        lblsad.Visible = True
        lblwink.Visible = True
        lblopen.Visible = True
        click2 = 1
    Else
        lblhappy.Visible = False
        lblsad.Visible = False
        lblwink.Visible = False
        lblopen.Visible = False
        frmchat.Height = 5310
        click2 = 0
    End If
End Sub

Private Sub lblopen_Click()
    txttyped.Text = txttyped.Text & " o :"
    lblhappy.Visible = False
    lblopen.Visible = False
    lblsad.Visible = False
    lblwink.Visible = False
    frmchat.Height = 5310
End Sub

Private Sub lblsad_Click()
    txttyped.Text = txttyped.Text & " ) :"
    lblhappy.Visible = False
    lblopen.Visible = False
    lblsad.Visible = False
    lblwink.Visible = False
    frmchat.Height = 5310
End Sub

Private Sub lblsend_Click()
    Dim typed As String
    typed = "Guest: " & txttyped.Text
    picconvo.ForeColor = &H800000
    c = c
    If c > 12 Then
        picconvo.top = picconvo.top - 390
        picconvo.top = picconvo.top
        picconvo.Height = picconvo.Height + 390
        picconvo.Height = picconvo.Height
        lblheader.top = lblheader.top + 390
        lblheader.top = lblheader.top
    End If
    picconvo.Print typed
    picconvo.Print
    c = c + 2
    Open App.Path & "\guestsaid.txt" For Output As #1
        Print #1, typed
        Close #1
    txttyped.Text = ""
    txttyped.SetFocus
    picconvo.ForeColor = &H8080&
    End Sub

Private Sub lblwink_Click()
    txttyped.Text = txttyped.Text & " ( ;"
    lblhappy.Visible = False
    lblopen.Visible = False
    lblsad.Visible = False
    lblwink.Visible = False
    frmchat.Height = 5310
End Sub

Private Sub Time2_Timer()
    If rickstatus = 0 Then
        lblheader.Caption = "MiM: Richard Motzko has gone off-line..."
        txttyped.Visible = False
    ElseIf rickstatus = 1 Then
        lblheader.Caption = "MiM: chat with Richard Motzko"
        txttyped.Visible = True
    End If
    clears = clears
    Open App.Path & "\ricksaid.txt" For Input As #1
        Input #1, recieved
        Close #1
    If recieved <> clears2 Then
        If c < 28 Then
            picconvo.ForeColor = &H8080&
            picconvo.Print recieved
            'picconvo.Print
            c = c + 2
            clears2 = recieved
        Else
            picconvo.ForeColor = &H8080&
            picconvo.top = picconvo.top - 195
            picconvo.top = picconvo.top
            picconvo.Height = picconvo.Height + 195
            picconvo.Height = picconvo.Height
            lblheader.top = lblheader.top + 195
            lblheader.top = lblheader.top
            picconvo.Print recieved
            'picconvo.Print
            c = c + 2
            clears2 = recieved
        End If
    Open App.Path & "\ricksaid.txt" For Output As #1
        Print #1, clears
        Close #1
    End If
End Sub
