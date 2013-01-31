VERSION 5.00
Begin VB.Form frmDrinks
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn
      Caption         =   "Return to Menu"
      BeginProperty Font
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   3
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton cmdOrder
      Caption         =   "Order"
      BeginProperty Font
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   2
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdLoadDrinks
      Caption         =   "What would you like to Drink?"
      BeginProperty Font
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   5880
      Width           =   2535
   End
   Begin VB.PictureBox gggg
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmDrinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoadDrinks_Click()
    Dim ffff As Long, Pos As Long
    'Load the Drinks file
    Open App.Path & "/Drinks.txt" For Input As #1
    ffff = 0
    Do Until EOF(1)
        ffff = ffff + 1
        Input #1, Drinks(ffff), Drinkcost(ffff)
    Loop
    Close #1
    'List the Drinks in the Picture Box
    gggg.Print "We Have:"
    gggg.Print "**********************"
    For Pos = 1 To ffff
    gggg.Print Drinks(Pos), FormatCurrency(Drinkcost(Pos))
    Next Pos
End Sub

Private Sub cmdOrder_Click()
    frmDrinks.Visible = False
    OrderForm.Visible = True

End Sub

Private Sub cmdReturn_Click()
    frmDrinks.Visible = False
    Menu.Visible = True
End Sub
