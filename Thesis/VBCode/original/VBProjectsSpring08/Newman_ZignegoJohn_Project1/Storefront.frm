VERSION 5.00
Begin VB.Form frmStorefront 
   BackColor       =   &H00FFFFFF&
   Caption         =   "store"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form2"
   Picture         =   "Storefront.frx":0000
   ScaleHeight     =   6600
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H000000C0&
      Caption         =   "Special"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdFoods 
      BackColor       =   &H000000C0&
      Caption         =   "Foods"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdEquipment 
      BackColor       =   &H000000C0&
      Caption         =   "Equipment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdApparel 
      BackColor       =   &H000000C0&
      Caption         =   "Apparel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtWelcome 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "Welcome to the Bazaar"
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmStorefront"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdApparel_Click()
frmStorefront.Hide
frmApparel.Show
End Sub

Private Sub cmdEquipment_Click()
frmStorefront.Hide
frmEquipment.Show
End Sub

Private Sub cmdFoods_Click()
frmStorefront.Hide
frmFood.Show
End Sub

Private Sub cmdSpecial_Click()
frmStorefront.Hide
frmSpecial.Show
End Sub
