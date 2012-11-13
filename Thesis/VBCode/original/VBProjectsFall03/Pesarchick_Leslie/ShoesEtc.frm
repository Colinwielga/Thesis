VERSION 5.00
Begin VB.Form frmShoesetc 
   BackColor       =   &H00404080&
   Caption         =   "Shoes Etc."
   ClientHeight    =   10215
   ClientLeft      =   1830
   ClientTop       =   345
   ClientWidth     =   10365
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   10215
   ScaleWidth      =   10365
   Visible         =   0   'False
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   16
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotalAccessories 
      Caption         =   "Total from Accessories"
      Height          =   615
      Left            =   1920
      TabIndex        =   15
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotalSkirts 
      Caption         =   "Total from Skirts"
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotalDresses 
      Caption         =   "Total from Dresses"
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotalShoes 
      Caption         =   "Total from Shoes"
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotalLeotards 
      Caption         =   "Total from Leotards"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotalUnitards 
      Caption         =   "Total from Unitards"
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Back to Main Menu"
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   3600
      ScaleHeight     =   6195
      ScaleWidth      =   6555
      TabIndex        =   6
      Top             =   2040
      Width           =   6615
   End
   Begin VB.CommandButton cmdSkirts 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Purchase Skirts"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdAccessories 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Purchase Accessories"
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdDresses 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Purchase Dresses"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdUnitards 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Purchase Unitards"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdLeotards 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Purchase Leotards"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdShoes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchase Shoes"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Helpful Hint: Purchase everything first, then click the totals buttons."
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   9720
      Width           =   2295
   End
   Begin VB.Label lblShoesetc 
      BackColor       =   &H00404080&
      Caption         =   "Click on what you want to Buy!"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmShoesetc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmShoesetc (Shoesetc.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy dance accessories
                    'the user can click on shoes, leotards, unitards
                    'dresses, skirts, and accessories
                    'allows the user to look at different types of products
                    'prints the totals for each category
                    'prints the overall total from everything purchased

Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.

Private Sub cmdAccessories_Click()
    frmAccessories.Show
    frmShoesetc.Hide
    frmAccessories.picResults.Cls
    frmAccessories.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmAccessories.picResults.Print "*************************************************************************************************"
End Sub

Private Sub cmdBack_Click()
    frmMain.Show
    frmShoesetc.Hide
End Sub

Private Sub cmdDresses_Click()
    frmDresses.Show
    frmShoesetc.Hide
    frmDresses.picResults.Cls
    frmDresses.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmDresses.picResults.Print "**************************************************************************************************"
End Sub

Private Sub cmdLeotards_Click()
    frmLeotards.Show
    frmShoesetc.Hide
    frmLeotards.picResults.Cls
    frmLeotards.picResults.Print "Item"; Tab(33); "Quantity"; Tab(45); "Price"
    frmLeotards.picResults.Print "***********************************************************************************************"
End Sub

Private Sub cmdShoes_Click()
    frmShoes.Show
    frmShoesetc.Hide
End Sub

Private Sub cmdSkirts_Click()
    frmSkirts.Show
    frmShoesetc.Hide
    frmSkirts.picResults.Cls
    frmSkirts.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmSkirts.picResults.Print "****************************************************************************************************"
End Sub

Private Sub cmdTotal_Click()
Total = TotalAccessories + TotalShoes + TotalDresses + TotalLeotards + TotalSkirts + TotalUnitards
picResults.Print "**********************************************************************************************************"
picResults.Print "Total", , FormatCurrency(Total)
End Sub

Private Sub cmdTotalAccessories_Click()
TotalAccessories = TotalAccessories + TotalAccessories2
picResults.Print "Total from Accessories", FormatCurrency(TotalAccessories)
End Sub

Private Sub cmdTotalDresses_Click()
picResults.Print "Total from Dresses", FormatCurrency(TotalDresses)
End Sub

Private Sub cmdTotalLeotards_Click()
picResults.Print "Total from Leotards", FormatCurrency(TotalLeotards)
End Sub

Private Sub cmdTotalShoes_Click()
TotalShoes = TotalBallet + TotalBallet2 + TotalTap + TotalJazz + TotalLyric
picResults.Print "Total from Shoes", FormatCurrency(TotalShoes)
End Sub

Private Sub cmdTotalSkirts_Click()
picResults.Print "Total from Skirts", FormatCurrency(TotalSkirts)
End Sub

Private Sub cmdTotalUnitards_Click()
picResults.Print "Total from Unitards", FormatCurrency(TotalUnitards)
End Sub

Private Sub cmdUnitards_Click()
    frmUnitards.Show
    frmShoesetc.Hide
    frmUnitards.picResults.Cls
    frmUnitards.picResults.Print "Item"; Tab(43); "Quantity"; Tab(55); "Price"
    frmUnitards.picResults.Print "***********************************************************************************************"
End Sub

