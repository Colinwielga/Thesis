VERSION 5.00
Begin VB.Form frmTarget 
   BackColor       =   &H000000C0&
   Caption         =   "frmTarget"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   Picture         =   "frmTarget.frx":0000
   ScaleHeight     =   10605
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Height          =   255
      Left            =   6840
      Picture         =   "frmTarget.frx":0FA5
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   240
      Picture         =   "frmTarget.frx":1767
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   10815
   End
   Begin VB.CommandButton cmdGame 
      Height          =   1095
      Left            =   240
      Picture         =   "frmTarget.frx":7FBD
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9360
      Width           =   2535
   End
   Begin VB.CommandButton cmdDVD 
      Height          =   1095
      Left            =   240
      Picture         =   "frmTarget.frx":B927
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton cmdFurniture 
      Height          =   1215
      Left            =   240
      Picture         =   "frmTarget.frx":F291
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdShipping 
      Height          =   4335
      Left            =   3240
      Picture         =   "frmTarget.frx":12BFB
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   7815
   End
   Begin VB.CommandButton cmdShoppingCart 
      Caption         =   "Shopping Cart"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdToys 
      Caption         =   "Toy Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6960
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdShoes 
      Caption         =   "Shoe Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdApparel 
      Caption         =   "Apparel Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdElectronics 
      Caption         =   "Electronics Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblSearch 
      BackColor       =   &H000000C0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblTarget 
      BackColor       =   &H000000C0&
      Caption         =   "Target"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2040
      TabIndex        =   6
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApparel_Click()
        frmTarget.Hide
        frmApparel.Show
End Sub

Private Sub cmdElectronics_Click()
        frmTarget.Hide
        frmElectronics.Show
End Sub

Private Sub cmdGo_Click()
        Dim CTR As Integer, Product(1 To 200) As String, Department(1 To 200) As String, Found As Boolean, Pos As Integer, Search As String
        Open App.Path & "\Search.txt" For Input As #1
        CTR = 0
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, Product(CTR), Department(CTR)
        Loop
        Close #1
        Found = False
        Pos = 1
        Search = txtSearch.Text
        Do While Not Found And Pos <= CTR
            If Search = Product(Pos) Then
                MsgBox "A " & Product(Pos) & " is in the " & Department(Pos) & " Department", , "Department"
                Found = True
            End If
            Pos = Pos + 1
        Loop
        If Found = False Then
            MsgBox "I'm Sorry, Target does not supply that product.", , "Error"
        End If
End Sub

Private Sub cmdHome_Click()
        frmTarget.Hide
        frmHome.Show
End Sub

Private Sub cmdShoes_Click()
        frmTarget.Hide
        frmShoes.Show
End Sub

Private Sub cmdShoppingCart_Click()
        frmTarget.Hide
        frmShoppingCart.Show
End Sub

Private Sub cmdToys_Click()
        frmTarget.Hide
        frmToys.Show
End Sub

Private Sub Form_Load()
        'Target, Corp.
        'frmTarget.frm
        'Mike Velin
        'March 23rd, 2009
        'Allow the user to click on different departments within the store, and search for specific items
End Sub
