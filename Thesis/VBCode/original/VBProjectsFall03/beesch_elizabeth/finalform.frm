VERSION 5.00
Begin VB.Form finalform 
   BackColor       =   &H00C000C0&
   Caption         =   "Total Cost"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   FillColor       =   &H00FF80FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton quit 
      Caption         =   "QUIT"
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton clear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton calculate 
      Caption         =   "SHOW TOTAL COST FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H80000012&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   240
      ScaleHeight     =   7155
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   960
      Width           =   6015
   End
End
Attribute VB_Name = "finalform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Screen Printing(Main1.vpb)
'Form Name : finalform(finalform.frm)
'Author: Elizabeth Beesch
'Date Written: October 27, 2003
'Purpose of Form: To print out the choices you have made or entered and print in
            'in an order form
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit

Private Sub Form_Load() ' when the form loads the following will happen
calculate.Enabled = True 'enables the calculate button
clear.Enabled = False 'disables the clear button
End Sub ' ends the commands of the button

Private Sub calculate_click() ' when the user clicks on this button the following will happen
Dim peritemtotal As Double
Dim subtotal As Single
Dim total As Double
total = 0
peritemtotal = 0
laborcost = (numberofItems * CTR * D) * 0.5 'calculates labor cost
total = totalscreens + apparelcost + namecost + numbercost + laborcost 'calculates total cost
peritemtotal = total / numberofItems ' calculates per item total
picResults.Print "Contact:"; Tab(5); contact 'prints contact information
picResults.Print "Phone Number:"; Tab(5); phonenumber 'prints phone number
picResults.Print "Job Order Name:"; Tab(5); joborder ' prints the job order name
picResults.Print "Date Needed By:"; Tab(5); dateneeded ' prints the date needed by
picResults.Print "Reorder:"; Tab(5); totalreorder ' prints if it is a reorder
picResults.Print "Number of Items:"; Tab(5); numberofItems ' prints the number of items
picResults.Print "Color of Apparel:"; Tab(5); finalcolor ' prints apparel color
picResults.Print "Color of Ink:"; Tab(5); inkcolor 'prints the color of ink that they choose
picResults.Print "Total Screens Cost:"; Tab(25); FormatCurrency(totalscreens, 2) 'prints total screen cost
picResults.Print "Total Apparel Cost:"; Tab(25); FormatCurrency(apparelcost, 2) ' prints total apparel cost
picResults.Print "Total Name Cost:"; Tab(25); FormatCurrency(namecost, 2) ' prints total name cost
picResults.Print "Total Number Cost:"; Tab(25); FormatCurrency(numbercost, 2) ' prints total number cost
picResults.Print "Total Labor Cost:"; Tab(25); FormatCurrency(laborcost, 2) ' prints labor cost
picResults.Print "_______________________________________________________________" ' prints a line
picResults.Print "Total Cost:"; FormatCurrency(total, 2) 'prints the total cost
picResults.Print
picResults.Print "Per Item Total:"; FormatCurrency(peritemtotal, 2) ' prints the total cost per item
calculate.Enabled = False ' disables the calculate button
clear.Enabled = True 'enables the clear button
MsgBox "Thanks For The Order", , "Thanks" 'message box appears
End Sub

Private Sub clear_Click() ' when the user clicks on this button the following will happen
picResults.Cls  'clears the picture box
calculate.Enabled = False 'disables the calculate button
clear.Enabled = False 'disables the clear button
End Sub ' ends the commands of the button

Private Sub quit_Click() ' when the user clicks on this button the following will happen
 End ' ends program
End Sub ' ends the commands of the button
