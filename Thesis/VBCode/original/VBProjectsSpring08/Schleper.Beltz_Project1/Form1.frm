VERSION 5.00
Begin VB.Form Tables 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   2400
      Picture         =   "Form1.frx":1619F
      ScaleHeight     =   2235
      ScaleWidth      =   6045
      TabIndex        =   8
      Top             =   0
      Width           =   6105
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H000000FF&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdInventory 
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   6
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdTable6 
      Caption         =   "Table 6"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6480
      Picture         =   "Form1.frx":2093E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdTable5 
      Caption         =   "Table 5"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4080
      Picture         =   "Form1.frx":20F9C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdTable4 
      Caption         =   "Table 4"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1680
      Picture         =   "Form1.frx":215FA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdTable3 
      Caption         =   "Table 3"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6480
      Picture         =   "Form1.frx":21C58
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdTable2 
      Caption         =   "Table 2"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4080
      Picture         =   "Form1.frx":222B6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdTable1 
      Caption         =   "Table 1"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1680
      Picture         =   "Form1.frx":22914
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
   End
End
Attribute VB_Name = "Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vinnie Joe's Pub
'Tables
'Vinnie Schleper, Joey Beltz
'3/13/08
'this is the second form you will see in the project.
'   From this form you can choose to look at inventory or proceed to keep track
'   of table orders.

Option Explicit
Private OldX As Integer
  Private OldY As Integer
  Private DragMode As Boolean
  Dim MoveMe As Boolean

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

Private Sub cmdInventory_Click()
' The "show" and "Hides" are used to move from one Table to the other.
Tables.Hide
Inventory.Show
End Sub

Private Sub cmdLogout_Click()
' This button is used to Logout.
Tables.Hide
Login.Show
End Sub

Private Sub cmdTable1_Click()
Tables.Hide
Table1.Show

End Sub

Private Sub cmdTable2_Click()
Tables.Hide
Table2.Show

End Sub

Private Sub cmdTable3_Click()
Tables.Hide
Table3.Show

End Sub

Private Sub cmdTable4_Click()
Tables.Hide
Table4.Show

End Sub

Private Sub cmdTable5_Click()
Tables.Hide
Table5.Show

End Sub

Private Sub cmdTable6_Click()
Tables.Hide
Table6.Show

End Sub

