VERSION 5.00
Begin VB.Form SubmitRecipe 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch10 
      BackColor       =   &H0000FF00&
      Caption         =   "Back to Main Form"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.PictureBox picIngDir 
      BackColor       =   &H00C0FFC0&
      Height          =   5775
      Left            =   3120
      ScaleHeight     =   5715
      ScaleWidth      =   7035
      TabIndex        =   2
      Top             =   360
      Width           =   7095
   End
   Begin VB.CommandButton cmdInputD 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter Directions"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdInputIn 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter Ingredients"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "SubmitRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSwitch10_Click()
Main.Show
SubmitRecipe.Hide
End Sub
