VERSION 5.00
Begin VB.Form frmMexican 
   BackColor       =   &H00808000&
   Caption         =   "Mexican"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMexican.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   9135
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   "Click here to see what you need for Beef Tacos"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdMReturn 
      BackColor       =   &H00C0C000&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdMQuit 
      BackColor       =   &H00C0C000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdMProcedures 
      BackColor       =   &H00C0C000&
      Caption         =   "Show Procedures"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   18
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
      Top             =   1920
      Width           =   2175
   End
   Begin VB.PictureBox picShowM 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   2760
      Picture         =   "frmMexican.frx":08CA
      ScaleHeight     =   4275
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
   End
   Begin VB.Label lblDish 
      BackColor       =   &H00808000&
      Caption         =   "Beef Tacos"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   480
      TabIndex        =   5
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   3585
      Left            =   1320
      Picture         =   "frmMexican.frx":13F8FC
      Top             =   5400
      Width           =   5490
   End
End
Attribute VB_Name = "frmMexican"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMProcedures_Click()

'Print Title
picShowM.Print "Mexican - Beef Tacos"
picShowM.Print "*******************************************************************"
picShowM.Print

'Display the procedures of Beef Tacos
picShowM.Print "1.Chop meat into 1/4-inch pieces. Heat a little oil."
picShowM.Print
picShowM.Print "2.Then cook the onion and garlic for about 2 minutes."
picShowM.Print
picShowM.Print "3.Stir in meat and brown it."
picShowM.Print
picShowM.Print "4.Add all to the beef mixture."
picShowM.Print
picShowM.Print "5.Stir in water and let mixture absorb it."
picShowM.Print
picShowM.Print "6.In a skillet, heat a couple tablespoons of oil."
picShowM.Print
picShowM.Print "7.Heat tortillas in the oil until soft."
picShowM.Print
picShowM.Print "8.Slice the avocado. While tortillas are hot, fill each with all mixtures."
picShowM.Print

End Sub


Private Sub cmdMQuit_Click()
End
End Sub

Private Sub cmdMReturn_Click()

'Return to Homepage
frmCountries.Show
frmMexican.Hide

End Sub

Private Sub cmdNext_Click()
groceryfile = "\Recipes\mexicanR.txt"

'Next Step
frmMexican.Hide
frmGroceryStore.Show


End Sub

