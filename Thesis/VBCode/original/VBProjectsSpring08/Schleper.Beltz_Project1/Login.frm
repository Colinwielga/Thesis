VERSION 5.00
Begin VB.Form Login 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   8835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   Picture         =   "Login.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
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
      Left            =   360
      TabIndex        =   14
      Top             =   7800
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   2280
      Picture         =   "Login.frx":1619F
      ScaleHeight     =   2235
      ScaleWidth      =   6045
      TabIndex        =   13
      Top             =   240
      Width           =   6105
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   11
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   10
      Left            =   4560
      TabIndex        =   10
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000000FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   6000
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   4560
      TabIndex        =   7
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   3120
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   6000
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   4560
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   6000
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vinnie Joe's Pub
'Login
'Vinnie Schleper, Joey Beltz
'3/13/08
'the purpose of this form is to gain access
'   to the rest of the program by use of passwords.
' The POS system is meant for a resturant who wants to keep track of
' inventory and to be able to enter orders and get totals.

' Our logo used was Photshopped from a different bar sign:
    'BarandPubSigns.com / Meakin Signs
        '1100 Davis Dr., Unit 18
        'Newmarket, ON L3Y 7V1
        'Canada
            'imgurl=http://www.barandpubsigns.com/images/Keswick_Pub_Sign.jpg&imgrefurl=http://www.barandpubsigns.com/pub_signs8.html&h=297&w=806&sz=87&hl=en&start=9&um=1&tbnid=cj_BqkNpNtfHKM:&tbnh=53&tbnw=143&prev=/images%





Option Explicit
  Private OldX As Integer
  Private OldY As Integer
  Private DragMode As Boolean
  Dim MoveMe As Boolean
' this code here is used for moving a window even when it doesn't have a border.
    ' Code was used from:
        ' Madboy on vbforums.com
            ' Date posted Nov. 30th 2003
                'http://www.vbforums.com/showthread.php?t=231152&goto=nextoldest
  Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     MoveMe = True
     OldX = X
     OldY = Y

 End Sub

 Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


     If MoveMe = True Then
         Me.Left = Me.Left + (X - OldX)
         Me.Top = Me.Top + (Y - OldY)
     End If

 End Sub

 Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


     Me.Left = Me.Left + (X - OldX)
     Me.Top = Me.Top + (Y - OldY)
     MoveMe = False

 End Sub



Private Sub cmd0_Click(Index As Integer)
'These pieces of code show that whatever is typed or inputed by
'   the text boxes or buttons are the same.
txtPassword.Text = txtPassword.Text & cmd0(Index).Caption
End Sub

Private Sub cmd1_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd1(Index).Caption
End Sub

Private Sub cmd2_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd2(Index).Caption
End Sub

Private Sub cmd3_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd3(Index).Caption
End Sub

Private Sub cmd4_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd4(Index).Caption
End Sub

Private Sub cmd5_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd5(Index).Caption
End Sub

Private Sub cmd6_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd6(Index).Caption
End Sub

Private Sub cmd7_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd7(Index).Caption
End Sub

Private Sub cmd8_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd8(Index).Caption
End Sub

Private Sub cmd9_Click(Index As Integer)
txtPassword.Text = txtPassword.Text & cmd9(Index).Caption
End Sub

Private Sub cmdClear_Click(Index As Integer)
txtPassword.Text = ""


End Sub

Private Sub cmdEnter_Click(Index As Integer)
'This "if" statement is used to decide whether the password entered
'   is true or false. If it is true then the Tables form is shown.
If txtPassword.Text = 8347 Then
    MsgBox "Welcome Vinnie!", , "Hello"
    Tables.Show
    Login.Hide
End If
If txtPassword.Text = 1387 Then
    MsgBox "Welcome Joey!", , "Hello"
    Tables.Show
    Login.Hide
End If
If txtPassword.Text <> 8347 And txtPassword.Text <> 1387 Then
    MsgBox "Invalid Password!", , "Sorry"
End If
    

End Sub

Private Sub cmdQuit_Click()
End
End Sub
