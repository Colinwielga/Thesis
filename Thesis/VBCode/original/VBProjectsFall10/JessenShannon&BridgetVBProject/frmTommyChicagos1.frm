VERSION 5.00
Begin VB.Form frm1Start 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13995
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   643
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   933
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdNewBGColor 
      Caption         =   "New Background Colors!"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11880
      TabIndex        =   4
      Top             =   8280
      Width           =   1335
   End
   Begin VB.PictureBox picPizzaPic 
      Height          =   5775
      Left            =   2760
      Picture         =   "frmTommyChicagos1.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   8115
      TabIndex        =   3
      Top             =   2160
      Width           =   8175
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5520
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   8160
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
   Begin VB.Label lblNames 
      BackColor       =   &H000000C0&
      Caption         =   "By: Bridget and Shannon Jessen"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   8760
      Width           =   3375
   End
   Begin VB.Label lblOnlineOrdering 
      BackColor       =   &H000000C0&
      Caption         =   "Online Ordering"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label lblTommyChicagos 
      BackColor       =   &H000000C0&
      Caption         =   "Tommy Chicago's Pizzeria "
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "frm1Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the first form of the program and it displays the name of the restaurant
'and what it does which is an online ordering program.
'There is a button called Menu which will take the user to the Menu page of the
'program. There is also a button that the user can click which changes the color
'of the background.


Private Sub cmdMenu_Click()
    
    'Initialize Variables
     pizzaListCtr = 0
    
    'loops through the array
    'this will list the pizza names chosen by the user
    'it will place them in the 0 column because that is where all the menu items are listed
    'pizzaList as an array are the names of the menu items
    
     Dim i As Integer
     For i = 0 To 50
        pizzaList(i, i) = "0"
     Next i
        
    'Switched the user from the first form to the second form
    frm2Menu.Show
    frm1Start.Hide
    
End Sub



Private Sub CmdNewBGColor_Click()

'This button will change the first background color to the next color and there
'are 5 different colors the user can see.
'backcolor is equal to one certain color
'the format with the ampersand is the technical title of the color
'colorcounter is incremented depending on how many times you click the button

colorCounter = colorCounter + 1

If colorCounter = 1 Then
    frm1Start.BackColor = &HFF0000
End If
If colorCounter = 2 Then
    frm1Start.BackColor = &HFF00&
End If
If colorCounter = 3 Then
    frm1Start.BackColor = &HC000C0
End If
If colorCounter = 4 Then
    frm1Start.BackColor = &H80FF&
End If
If colorCounter = 5 Then
    frm1Start.BackColor = &HC000&
End If
If colorCounter = 6 Then
    frm1Start.BackColor = &HC0&
End If
    
    
End Sub
