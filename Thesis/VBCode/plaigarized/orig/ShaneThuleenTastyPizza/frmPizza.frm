VERSION 5.00
Begin VB.Form frmPizza 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Back to Menu"
      Height          =   975
      Left            =   5880
      TabIndex        =   3
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order"
      Height          =   975
      Left            =   3120
      TabIndex        =   2
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton cmdLoadPizza 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Here is what we have"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   2775
   End
   Begin VB.PictureBox picPizza 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6315
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmPizza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoadPizza_Click()
    Dim Pizza(1 To 100) As String, Pcost(1 To 100)
    Dim CTR As Integer, Pos As Integer
    'Load the Pizza file
    Open App.Path & "/Pizza.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Pizza(CTR), Pcost(CTR)
    Loop
    Close #1
    'List the Pizza in the Picture Box
    picPizza.Print "We Have:"
    picPizza.Print "**********************"
    For Pos = 1 To CTR
        picPizza.Print Pizza(Pos), FormatCurrency(Pcost(Pos))
    Next Pos
End Sub

Private Sub cmdOrder_Click()
    frmPizza.Visible = False
    OrderForm.Visible = True
End Sub

Private Sub cmdReturn_Click()
    frmPizza.Visible = False
    Menu.Visible = True
End Sub
