VERSION 5.00
Begin VB.Form frmRecipe
   BackColor       =   &H0000FFFF&
   Caption         =   "Change Recipe"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrice
      Caption         =   "Change my price"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   8880
      Width           =   1575
   End
   Begin VB.TextBox txtPrice
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Text            =   "1"
      Top             =   8400
      Width           =   1575
   End
   Begin VB.PictureBox picRecipe
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   2595
      TabIndex        =   9
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Return"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange
      Caption         =   "Change my recipe"
      Height          =   1335
      Left            =   2880
      TabIndex        =   7
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtIce
      Height          =   1095
      Left            =   2160
      TabIndex        =   6
      Text            =   "1"
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtSugar
      Height          =   1095
      Left            =   2160
      TabIndex        =   5
      Text            =   "1"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtLemons
      Height          =   1095
      Left            =   2160
      TabIndex        =   4
      Text            =   "3"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblPrice
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter the price of your lemonade. Remember, expensive lemonade will not sell as well, unless it's really good."
      BeginProperty Font
         Name            =   "@Gulim"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Label lblIce
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter the number of trays of ice to be used per pitcher."
      BeginProperty Font
         Name            =   "@Gulim"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblSugar
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter the amount of sugar to be used per pitcher."
      BeginProperty Font
         Name            =   "@Gulim"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblLemons
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter the number of lemons to be used per pitcher."
      BeginProperty Font
         Name            =   "@Gulim"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblRecipe
      BackColor       =   &H00FFFF00&
      Caption         =   $"frmRecipe.frx":0000
      BeginProperty Font
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Dim aaaa as Integer = 1
    frmRecipe.Hide
    frmMainScreen.Show
End Sub

Private Sub cmdChange_Click()
    RecipeL = txtLemons
    RecipeS = txtSugar
    RecipeI = txtIce
    If RecipeL <= 0 Or RecipeS <= 0 Or RecipeI <= 0 Then
        MsgBox "AHHH! Your lemonade is freaky; make sure you actually have SOME lemons, sugar, and ice!"
    ElseIf RecipeL + RecipeS + RecipeI > 25 Then
        MsgBox "Hey, one pitcher can only hold 25 units of stuff -- you've got to adjust your recipe!"
    Else
        picRecipe.Cls
        picRecipe.Print "My current recipe for one pitcher is..."
        picRecipe.Print "Lemons:", RecipeL
        picRecipe.Print "Sugar:", RecipeS
        picRecipe.Print "Ice:", RecipeI
    End If
End Sub

Private Sub cmdPrice_Click()
    charged = txtPrice
    If charged <= 0 Then
        MsgBox "Wow. Try actually charging your customers SOMETHING. Is your lemonade really that bad?"
    End If
End Sub
