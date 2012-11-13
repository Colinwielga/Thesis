VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H00000080&
   Caption         =   "Food Pyramid"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back to Main Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdidealweight 
      Caption         =   "Go to Ideal Weight Calculation"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      TabIndex        =   7
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "X"
      Height          =   495
      Left            =   11040
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdprotein 
      Caption         =   "Proteins"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   21
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      Picture         =   "FoodPyramid.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmddairy 
      Caption         =   "Dairy"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Picture         =   "FoodPyramid.frx":103B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdveggie 
      Caption         =   "Vegetables"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      Picture         =   "FoodPyramid.frx":1F2E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdfriut 
      Caption         =   "Fruits"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      Picture         =   "FoodPyramid.frx":A04E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdgrain 
      Caption         =   "Grains"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      Picture         =   "FoodPyramid.frx":AC56
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   7815
   End
   Begin VB.CommandButton cmdfat 
      Caption         =   "Oils"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      Picture         =   "FoodPyramid.frx":47D7C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'frm1
'Ben Morris
'March 21
'Homepage for the food pyramid

Option Explicit

Private Sub cmddairy_Click()
    frm1.Hide
    frm2.Hide
    frm3.Hide
    frm4.Hide
    frm5.Hide
    frm6.Show
    frm7.Hide
    frm8.Hide
    ' this hides all other forms and makes the dairy form visible
End Sub

Private Sub cmdfat_Click()
    frm1.Hide
    frm2.Hide
    frm3.Hide
    frm4.Hide
    frm5.Hide
    frm6.Hide
    frm7.Hide
    frm8.Show
    ' this hides all other forms and makes the oils form visible
End Sub

Private Sub cmdfriut_Click()
     frm1.Hide
    frm2.Hide
    frm3.Hide
    frm4.Show
    frm5.Hide
    frm6.Hide
    frm7.Hide
    frm8.Hide
    ' this hides all other forms and makes the fruit form visible
End Sub

Private Sub cmdgrain_Click()
    frm1.Hide
    frm2.Hide
    frm3.Show
    frm4.Hide
    frm5.Hide
    frm6.Hide
    frm7.Hide
    frm8.Hide
    ' this hides all other forms and makes the grain form visible
End Sub

Private Sub cmdidealweight_Click()
    frm1.Hide
    frm2.Show
    ' this hides the pyramid form and makes the weight calculator visible
End Sub

Private Sub cmdprotein_Click()
    frm1.Hide
    frm2.Hide
    frm3.Hide
    frm4.Hide
    frm5.Hide
    frm6.Hide
    frm7.Show
    frm8.Hide
    ' this hides all other forms and makes the protien form visible
End Sub

Private Sub cmdquit_Click()
    End
    'this ends the program
End Sub

Private Sub cmdveggie_Click()
     frm1.Hide
    frm2.Hide
    frm3.Hide
    frm4.Hide
    frm5.Show
    frm6.Hide
    frm7.Hide
    frm8.Hide
    'this hides all other forms and makes the veggie form visible
End Sub

Private Sub Command1_Click()
frm1.Hide
frmMainpage.Show
End Sub
