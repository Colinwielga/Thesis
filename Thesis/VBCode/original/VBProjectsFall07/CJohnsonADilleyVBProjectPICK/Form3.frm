VERSION 5.00
Begin VB.Form frmPick 
   BackColor       =   &H0000FF00&
   Caption         =   "Drinks"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form3"
   ScaleHeight     =   8760
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcalculate 
      BackColor       =   &H00800080&
      Caption         =   "Calculate my BAC!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7200
      MaskColor       =   &H00800080&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   10200
      ScaleHeight     =   6075
      ScaleWidth      =   4755
      TabIndex        =   12
      Top             =   2040
      Width           =   4815
   End
   Begin VB.CommandButton cmdShot 
      Height          =   2175
      Left            =   240
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton cmdWine 
      Height          =   2535
      Left            =   3840
      Picture         =   "Form3.frx":078D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdFruit 
      Height          =   2295
      Left            =   3600
      Picture         =   "Form3.frx":100C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdBeer 
      Height          =   2775
      Left            =   6600
      Picture         =   "Form3.frx":1A18
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblConsume 
      BackColor       =   &H008080FF&
      Caption         =   "You Have Consumed:"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   13
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblBACL 
      BackColor       =   &H00C0FFC0&
      Caption         =   $"Form3.frx":424B
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblWine 
      BackColor       =   &H00808080&
      Caption         =   "5 oz of Wine (~11% ABV)"
      BeginProperty Font 
         Name            =   "Minion Pro Med"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblCocktail 
      BackColor       =   &H00808080&
      Caption         =   "CoCkTaIl-1.5 oz of distilled Liquor (11% ABV)"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4200
      TabIndex        =   9
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label lblShot 
      BackColor       =   &H00808080&
      Caption         =   "1 oz Shot (~45%ABV)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label lblBeer 
      BackColor       =   &H00808080&
      Caption         =   "12 oz Beer (5% ABV)"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblPoison 
      BackColor       =   &H0000FF00&
      Caption         =   "Poison"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      TabIndex        =   2
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblYour 
      BackColor       =   &H0000FF00&
      Caption         =   "YOUR"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblPick 
      BackColor       =   &H0000FF00&
      Caption         =   " Pick"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'dim our variables
 Dim quantity As Single, quan As Single, tity As Single, quantit As Single
 

Private Sub cmdBeer_Click()
'ask the user how many beers they have had via inputbox
quantity = InputBox("How many have you had?")
'the sum is the alcohol equivalent content in one beer multiplied by the quantity
sum1 = (1.2 * quantity)
'in the results box we list the quantity and item
picResults.Print quantity; " 12 oz. beer(s)"
End Sub

Private Sub cmdFruit_Click()
'ask the user how many mixed drinks they have had via inputbox
quan = InputBox("How many have you had?")
    'in the results box we list the quantity and item
picResults.Print quan; " Mixed Drink(s)-1.5oz of liquor"
'the sum is the alcohol equivalent content in one drink multiplied by the quantity
sum2 = (0.33 * quan)
End Sub

Private Sub cmdShot_Click()
'ask the user how many shots they have had via inputbox
tity = InputBox("How many have you had?")
    'in the results box we list the quantity and item
picResults.Print tity; " 1 oz. shot(s)"
'the sum is the alcohol equivalent content in one drink multiplied by the quantity
sum3 = (0.9 * tity)
End Sub

Private Sub cmdWine_Click()
'ask the user how many glasses they have had via inputbox
quantit = InputBox("How many have you had?")
    'in the results box we list the quantity and item
picResults.Print quantit; " 5 oz. glass(es)of wine"
'the sum is the alcohol equivalent content in one drink multiplied by the quantity
sum4 = (1.1 * quantit)
End Sub

Private Sub cmdcalculate_Click()
'we add all the sums together to find actual consumption total
sum = sum1 + sum2 + sum3 + sum4
'we hide the current form and show next form
frmPick.Hide
frmInfo.Show

End Sub

