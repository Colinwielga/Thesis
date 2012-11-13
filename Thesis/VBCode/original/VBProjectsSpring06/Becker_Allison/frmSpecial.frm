VERSION 5.00
Begin VB.Form frmSpecial 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   4920
      Picture         =   "frmSpecial.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   5760
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2880
      Picture         =   "frmSpecial.frx":B17A
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   5760
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   600
      Picture         =   "frmSpecial.frx":152B4
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H008080FF&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   6000
      ScaleHeight     =   1995
      ScaleWidth      =   3795
      TabIndex        =   5
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txtTotal 
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdTotals 
      BackColor       =   &H008080FF&
      Caption         =   "Calculate Total with Discount"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Designed by Allison Becker"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSpecial.frx":21FC6
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1815
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   8295
   End
   Begin VB.Label lblSpecial 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "March Special"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1455
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   8415
   End
   Begin VB.Label lblInput 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Input # of Flowers you want to Purchase"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   2775
   End
End
Attribute VB_Name = "frmSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Flowers For U! (FlowerShop.vbp)
'Form Name: (frmSpecial)
'Author: Allison Becker
'Date Written: 3/23/06
'Objective: The objective of this form to show the user our monthly special. It also
'allows them to input the number of flowers they would like to purchase and see a total
'with the prices dicounted.
Option Explicit

Private Sub cmdBack_Click()
    frmSpecial.Hide
    frmFlowerShop.Show
End Sub

Private Sub cmdTotals_Click()
    Dim Amount As Integer
    Dim Total As Single
    'If then statement
    Amount = txtTotal.Text 'taking input from a text box
    picResults.Cls
    If Amount > 5 Then
        Total = Amount * 1.5
    Else
        Total = Amount * 3
    End If
    picResults.Print "Total is", FormatCurrency(Total)
End Sub



