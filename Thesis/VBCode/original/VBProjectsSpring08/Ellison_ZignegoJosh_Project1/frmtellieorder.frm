VERSION 5.00
Begin VB.Form frmtellieorder 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Order Something"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox pick 
      Height          =   315
      ItemData        =   "frmtellieorder.frx":0000
      Left            =   3720
      List            =   "frmtellieorder.frx":0002
      TabIndex        =   11
      Text            =   "Choose a size"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00F4C6A4&
      Caption         =   "Continue on your tour de st. joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H00F4C6A4&
      Caption         =   "Return to Tellies main page"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00F4C6A4&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdtotal 
      BackColor       =   &H00F4C6A4&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdsandwich 
      BackColor       =   &H008DD9FA&
      Caption         =   "Sandwich"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdsticks 
      BackColor       =   &H008DD9FA&
      Caption         =   "Bread Sticks"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdsoda 
      BackColor       =   &H008DD9FA&
      Caption         =   "Soda"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdsalad 
      BackColor       =   &H008DD9FA&
      Caption         =   "Salad"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdpizza 
      BackColor       =   &H008DD9FA&
      Caption         =   "Pizza"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   6120
      ScaleHeight     =   4515
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label liblordering 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on what you want to order"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   14160
      Left            =   -4680
      Picture         =   "frmtellieorder.frx":0004
      Top             =   -2520
      Width           =   20295
   End
End
Attribute VB_Name = "frmtellieorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningtotal As Single

 
    'Project name:  Tour De St. Joe
    'Form:  frmtellieorder, "Order"
    'Author:  Brooke and Josh
    'Date:  3/11/08
    'Objective: Ask the user to select what types of food they would like to "buy" and then adds up the total bill.

Private Sub cmdclear_Click()

    picoutput.Cls

    runningtotal = 0

End Sub

Private Sub cmdleave_Click()

    frmjoetown.Show
    frmtellieorder.Hide

End Sub

Private Sub cmdpizza_Click()
    'declair the variables
    
    Dim pizza As Single
    Dim pizza2 As Single
    Dim pizza3 As Single
    Dim myarray As String
    Dim Size As String

    pizza = 5.75
    pizza2 = 8.25
    pizza3 = 12.5

    Size = pick.Text
       
    If Size = "Small" Then
            picoutput.Print "Pizza ", FormatCurrency(pizza)
            runningtotal = pizza + runningtotal
    ElseIf Size = "Medium" Then
            picoutput.Print "Pizza ", FormatCurrency(pizza2)
            runningtotal = pizza2 + runningtotal
    ElseIf Size = "Huge" Then
            picoutput.Print "Pizza ", FormatCurrency(pizza3)
            runningtotal = pizza3 + runningtotal
    End If
    
End Sub

Private Sub cmdreturn_Click()

    frmtellies.Show
    frmtellieorder.Hide

End Sub

Private Sub cmdsalad_Click()
    
    Dim salad As Single
    salad = 3.5
    runningtotal = salad + runningtotal
    
    picoutput.Print "Salad", FormatCurrency(salad)

End Sub

Private Sub cmdsandwich_Click()
    
    Dim sand As Single
    sand = 7.75
    runningtotal = sand + runningtotal
    
    picoutput.Print "Sandwich ", FormatCurrency(sand)
    
End Sub

Private Sub cmdsoda_Click()
    
    Dim soda As Single
    soda = 1.75
    runningtotal = soda + runningtotal
    
    picoutput.Print "Soda ", FormatCurrency(soda)

End Sub

Private Sub cmdsticks_Click()
    
    Dim sticks As Single
    sticks = 5.5
    runningtotal = sticks + runningtotal
    
    picoutput.Print "Bread Sticks ", FormatCurrency(sticks)

End Sub

Private Sub cmdtotal_Click()

picoutput.Print "--------------------"

    picoutput.Print "Subtotal", FormatCurrency(runningtotal)
    
    Dim tax As Single
    tax = runningtotal * 0.07
    
    picoutput.Print "Tax", FormatCurrency(tax)
    
    Dim total As Single
    total = runningtotal + tax
    
    picoutput.Print "Total", FormatCurrency(total)

End Sub

Private Sub Form_Load()

    pick.AddItem "Small"
    pick.AddItem "Medium"
    pick.AddItem "Huge"

End Sub
