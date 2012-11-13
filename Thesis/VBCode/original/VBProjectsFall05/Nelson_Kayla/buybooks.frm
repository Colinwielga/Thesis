VERSION 5.00
Begin VB.Form frmbuybooks 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Buy Books; Project by Kayla Nelson"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   ForeColor       =   &H00800080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FF80FF&
      Caption         =   "Clear Total"
      Height          =   1335
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdtotal 
      BackColor       =   &H00FF80FF&
      Caption         =   "Calculate Total"
      Height          =   1335
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.PictureBox picrunningtotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   5
      Top             =   5160
      Width           =   4095
   End
   Begin VB.PictureBox piccity 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   7680
      Picture         =   "buybooks.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   1200
      Width           =   1530
   End
   Begin VB.PictureBox pickite 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   5520
      Picture         =   "buybooks.frx":1CA1
      ScaleHeight     =   2460
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   1320
      Width           =   1605
   End
   Begin VB.PictureBox picfirst 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3120
      Picture         =   "buybooks.frx":A51A
      ScaleHeight     =   2700
      ScaleWidth      =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1830
   End
   Begin VB.PictureBox picmillion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   600
      Picture         =   "buybooks.frx":C188
      ScaleHeight     =   2970
      ScaleWidth      =   1950
      TabIndex        =   1
      Top             =   840
      Width           =   1980
   End
   Begin VB.CommandButton cmdreturnmm 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Please Click on the Book to Add to Total:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   10095
   End
   Begin VB.Label lblprice4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "$25.95"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblprice3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "$14.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblprice2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "$24.95"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblprice1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "$14.95"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "frmbuybooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Kayla's Book Club (MainMenu.vbp)
'Form Name: Buy Books (frmbuybooks.frm)
'Author: Kayla Nelson
'Date: 10-27-05
'purpose of the form: This form allows the user to click on the picture of the book to bring up the price in a picture box.  You can then select a total button to see your total.  You may also clear your running total or return to the main menu form.

Dim Subtotal As Double
Option Explicit

Private Sub cmdclear_Click()
    picrunningtotal.Cls 'Clears the running total being made within the picture box.
    Subtotal = 0 'Sets the subtotal back to 0
End Sub

Private Sub cmdreturnmm_Click() ' This closes the Buy books form and goes back to the main menu page.
    frmbuybooks.Hide
    frmmainmenu.Show
End Sub



Private Sub cmdtotal_Click()
    picrunningtotal.Print "******************"
    picrunningtotal.Print FormatCurrency(Subtotal) 'This prints the users subtotal for all the books they selected.
End Sub



Private Sub piccity_Click()
    picrunningtotal.Print "$25.95" 'prints price in picturebox
    Subtotal = Subtotal + 25.95 'adds price to running subtotal
End Sub

Private Sub picfirst_Click()
    picrunningtotal.Print "$24.95" 'prints price in picturebox
    Subtotal = Subtotal + 24.95 'adds price to running subtotal
End Sub

Private Sub pickite_Click()
    picrunningtotal.Print "$14.00" 'prints price in picturebox
    Subtotal = Subtotal + 14 'adds price to running subtotal
End Sub

Private Sub picmillion_Click()
    picrunningtotal.Print "$14.95" 'prints price in picturebox
    Subtotal = Subtotal + 14.95 'adds price to running subtotal
End Sub
